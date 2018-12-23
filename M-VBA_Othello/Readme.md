# 행정병의 VBA 도전기#2
#### Othello 게임 개발기

---
## 개요

군대는 컴퓨터의 사양을 떠나 두 가지 분류로 나뉜다.

* 인터넷 망에 연결된 PC (사이버지식정보방 포함)
* 행정 PC(인터넷에 연결되어 있지 않는 말그대로 행정업무용 PC)

현역 행정병의 경우 전자보다는 후자를 더 많이 다루더라....

이건 부대마다 차이가 있다.

어쨋든 행정병이 행정 PC를 사용하여 할 수 있는 것이 없다.

하지만 Excel이 설치되어 있는 행정 PC 이면 이야기가 달라진다.

필자는 Excel VBA를 사용하여 오목, 오셸로 등

다양한 게임을 만들었다.

오늘은 내가 만든 게임 중 하나인 오셸로를 소개 해볼까 한다.

---

## 설계

게임 룰은 [여기](https://namu.wiki/w/%EC%98%A4%EB%8D%B8%EB%A1%9C(%EB%B3%B4%EB%93%9C%20%EA%B2%8C%EC%9E%84))에 잘 설명되어 있으니 여길 보도록 하자.

게임 구현을 위해 필요한 것은 아래와 같다.

* 8 X 8 = 64 칸의 공간
* 바둑돌 디자인(필자는 PPT를 사용하여 만들었다.)
* 돌 세팅(정 중앙에 흑 백 두개의 돌을 교차)
* 둘 수 있는 공간을 알려주는 기능(둘 수 없을 경우 알림)

뭐 우리 끼리 할 것이기 그렇게 많은 기능이 필요하지는 않는다.

그래도 필요한 기능이 있어서 새로운 기능이 추가된 경우 계속 수정할 것이다.

만들어 놓은 돌맹이 PPT도 공유하고싶었으나 파일을 찾을 수 없다..ㅠ

나중에 오목 만드실 분들 안에있는 이미지 갖다 쓰세요

---

## 기능 구현

기능 구현을 위해 먼저 3개의 개체를 구현했습니다.

먼저 게임 시작 시 간단하게 기능을 알려 줘야겟죠?

MsgBox를 이용해서 구현했습니다.

```vba
Private Sub Workbook_Open()
    Application.OnKey "{F5}", "GameModule.StartGame"
    Application.OnKey "{Backspace}", "GameModule.Undo"
    MsgBox "F5 : 게임 시작(재시작)" & vbNewLine & _
           "Backspace : 무르기" & vbNewLine & _
           "시트 수정 시 오작동을 일으킴.", vbInformation, ""
End Sub
```

그리고 그래픽 부분을 위한 그래픽 개체를 만들어 줍니다.

```vba
Private Const STONE_WHITE As String = "STONE_WHITE"
Private Const STONE_BLACK As String = "STONE_BLACK"

Private selectSheet As Boolean

Public Sub ClearClickableArea()
    With Sheet1.Range("A1:H8").Interior
        .Pattern = xlNone
    End With
End Sub

'//선택가능한 영역을 회색으로 칠함
Public Sub PaintClickableArea(y As Integer, x As Integer)
    With Sheet1.Cells(y, x).Interior
        .Pattern = xlSolid
        .Color = RGB(230, 230, 230)
    End With
End Sub

'//마우스가 아닌 코드로 클릭 시 구분위해 둔 함수
Public Function IsCodeSelectSheet() As Boolean
    IsCodeSelectSheet = selectSheet
End Function

Public Function GetNewStone(st As StoneTypes) As Shape

    Dim src As Shape
    Set src = GetBaseStone(st)

    src.Copy
    Sheet1.Paste

    Set GetNewStone = Sheet1.Shapes(Sheet1.Shapes.Count)

End Function

'//새로 비치한 돌의 그래픽을 시트에 추가
Public Sub CreateStoneGraphic(y As Integer, x As Integer, st As StoneTypes)

    Dim stShp As Shape
    Set stShp = GraphicModule.GetNewStone(st)

    With stShp
        .Name = y & "a" & x
        .Left = Sheet1.Cells(y, x).Left + 2
        .Top = Sheet1.Cells(y, x).Top + 2
    End With

    selectSheet = True
    Sheet1.Range("I1").Select
    selectSheet = False

End Sub

'//돌 뒤집기용. 따로 도형을 삭제해서 추가하는 것이 아니라 서식 복사
Public Sub ChangeStoneGraphic(y As Integer, x As Integer, newType As StoneTypes)

    Dim src As Shape, tg As Shape

    Set src = GetBaseStone(newType)
    Set tg = Sheet1.Shapes(y & "a" & x)

    src.PickUp
    tg.Apply

End Sub

Private Function GetBaseStone(st As StoneTypes) As Shape

    If st = stBlack Then
        Set GetBaseStone = Sheet2.Shapes(STONE_BLACK)
    Else
        Set GetBaseStone = Sheet2.Shapes(STONE_WHITE)
    End If

End Function

Public Sub DeleteStoneGraphics()

    Dim i As Integer

    Do While Sheet1.Shapes.Count > 0
        Sheet1.Shapes(Sheet1.Shapes.Count).Delete
    Loop

End Sub
```

마지막으로 게임 Rule에 맞춰 게임 모듈을 만들어줍니다.

```vba
Private Const SIZE As Integer = 9

Public Enum StoneTypes
    stEmpty = 0
    stBlack = 1
    stWhite = 2
End Enum

Private xOffsets As Variant
Private yOffsets As Variant

Private curStType As StoneTypes

Private map(9, 9) As StoneTypes
Private lastMap(9, 9) As StoneTypes '//무르기 용

Private reverseList As Collection

Private isSaveReverseStone As Boolean
Private isDisplayClickableArea As Boolean

Private isRunning As Boolean
Private canUndo As Boolean

Public Sub StartGame()

    Dim i As Integer, r As Integer

    isRunning = True

    '//시트 그래픽 초기화
    Call GraphicModule.ClearClickableArea
    Call GraphicModule.DeleteStoneGraphics

    '//게임 변수 초기화
    For i = 0 To 8
        For j = 0 To 8
            map(i, j) = stEmpty
        Next
    Next

    Set reverseList = New Collection

    xOffsets = Array(-1, 0, 1)
    yOffsets = Array(-1, 0, 1)

    selectSheet = False
    curStType = stBlack

    isDisplayClickableArea = True
    isSaveReverseStone = True

    '//초기 돌 배치
    Call ForcePutStone(4, 4, stWhite)
    Call ForcePutStone(4, 5, stBlack)
    Call ForcePutStone(5, 5, stWhite)
    Call ForcePutStone(5, 4, stBlack)

    '//클릭 할 수 있는 영역 표시
    Call DisplayClickableCell

    '//무르기를 위한 Map Copy
    Call CopyCurrentMap
    canUndo = False

End Sub

'//선택할 수 있는 셀 개수 반환(예정)
Private Sub DisplayClickableCell()

    Dim y As Integer, x As Integer
    Dim cnt As Integer

    If isDisplayClickableArea = False Then Exit Sub

    isSaveReverseStone = False

    Call GraphicModule.ClearClickableArea

    For y = 1 To 8
        For x = 1 To 8
            If map(y, x) = stEmpty Then
                If PutStone(y, x) Then
                    Call GraphicModule.PaintClickableArea(y, x)
                End If
                map(y, x) = stEmpty
            End If
        Next
    Next

    isSaveReverseStone = True

End Sub

'//시트 클릭 시
Public Sub HandleClickMap(y As Integer, x As Integer)

    If isRunning = False Then Call StartGame

    '//무르기를 위한 배열 복사
    Call CopyCurrentMap

    '//돌 놓기 성공
    If PutStone(y, x) = True Then
        '//시트에 클릭한 지점 돌 그래픽 생성
        Call CreateStoneGraphic(y, x, curStType)
        '//돌 뒤집기
        Call ReverseStone
        '//선택가능한 셀 표시
        Call DisplayClickableCell
    Else
        MsgBox "그 곳에는 놓을 수 없습니다!", vbCritical, ""
    End If

End Sub

'//무르기를 위한 배열 복사
Private Sub CopyCurrentMap()

    Dim i As Integer, j As Integer

    canUndo = True

    For i = 0 To 9
        For j = 0 To 9
            lastMap(i, j) = map(i, j)
        Next
    Next

End Sub

'//무르기
Public Sub Undo()

    Dim y As Integer, x As Integer

    If canUndo = False Then
        MsgBox "무르기를 더 이상 할 수 없습니다.", vbCritical, ""
        Exit Sub
    End If

    '//시트에 표시된 모든 그래픽 제거
    Call GraphicModule.DeleteStoneGraphics
    Call GraphicModule.ClearClickableArea

    For y = 0 To 9
        For x = 0 To 9
            map(y, x) = lastMap(y, x)
            If (map(y, x) <> stEmpty) Then
                Call GraphicModule.CreateStoneGraphic(y, x, map(y, x))
            End If
        Next
    Next

    curStType = IIf(curStType = stBlack, stWhite, stBlack)

    Call DisplayClickableCell

    canUndo = False

End Sub

'//돌 뒤집기
Private Sub ReverseStone()

    Dim y As Integer, x As Integer

    Do While reverseList.Count > 0

        y = CInt(reverseList.item(reverseList.Count)(0))
        x = CInt(reverseList.item(reverseList.Count)(1))
        map(y, x) = curStType

        Call GraphicModule.ChangeStoneGraphic(y, x, curStType)

        reverseList.Remove reverseList.Count

    Loop

    curStType = IIf(curStType = stBlack, stWhite, stBlack)

End Sub

Private Function PutStone(y As Integer, x As Integer) As Boolean

    Dim r As Integer, c As Integer
    Dim xo As Integer, yo As Integer
    Dim ret As Boolean

    map(y, x) = curStType

    '//클릭한 지점으로 부터 8방향 탐색
    For r = 0 To 2
        For c = 0 To 2
            If r <> 1 Or c <> 1 Then
                xo = CInt(xOffsets(c))
                yo = CInt(yOffsets(r))
                ret = ret Or FindReverseStones(y + yo, x + xo, yo, xo, curStType)
            End If
        Next
    Next

    '//뒤집을 돌이 없는 경우 돌 놓기 취소
    If ret = False Then
        map(y, x) = stEmpty
    End If

    PutStone = ret

End Function

Private Function FindReverseStones(y As Integer, x As Integer, yo As Integer, xo As Integer, st As StoneTypes) As Boolean

    Dim ret As Boolean

    If y < 1 Or x < 1 Or x > SIZE Or y > SIZE Then FindReverseStones = False: Exit Function
    If map(y, x) = stEmpty Then FindReverseStones = False: Exit Function

    If map(y, x) = st Then
        If map(y - yo, x - xo) <> stEmpty And map(y - yo, x - xo) <> st Then FindReverseStones = True
        Exit Function
    End If

    If FindReverseStones(y + yo, x + xo, yo, xo, st) = True Then
        If isSaveReverseStone Then reverseList.Add Array(y, x)
        FindReverseStones = True
    End If

End Function

'//뒤집을 돌 검사안하고 그냥 강제로 돌 놓는 함수
Private Sub ForcePutStone(y As Integer, x As Integer, st As StoneTypes)
    map(y, x) = st
    Call GraphicModule.CreateStoneGraphic(y, x, st)
End Sub
```

이제 시트를 통해 정상적으로 게임이 될 수 있도록 해야겠죠?

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error Resume Next

    If GraphicModule.IsCodeSelectSheet = True Then Exit Sub
    If Target.Count > 1 Then Exit Sub
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    Call GameModule.HandleClickMap(Target.Row, Target.Column)

End Sub
```
플레이는 알아서^^
