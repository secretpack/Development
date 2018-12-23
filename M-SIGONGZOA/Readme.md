# 행정병의 VBA 도전기#1
#### 시공의 폭풍 구현

---
## 개요

이걸 처음부터 노리고 만든 것은 아니였다.

VBA 도움말을 읽으며 shapes 함수를 잘 사용하면 객체를 빙글빙글 돌릴 수 있겠다는 재미있는 발상에 착안하였고

그러던 중 후임이 시공의 폭풍을 만들어 달라고 요청했다.

그래서 PowerPoint를 사용하여 끝없는 노가다로 마크를 구현하고

VBA 매크로를 사용하여 이를 구현했다.

먼저 Module1 을 아래와 같이 구현한다.

VBA를 익힐때 가장 불편했던 것은 += 와 같은 연산자가 없다는 것이다.

예를들어 a = a + 10을 할때 C언어에서는 다음과 같이 표현할 수 있다.

```c
int a;

a = a + 10;
a += 10;
```

하지만 VBA에서는 그게 안된다... 무조건 전자 처럼 써야한다.

코드를 구현해보자

```vba
Public isRotating As Boolean

Public Sub RotateHOS()

  Dim val As Interger
  Dim inVal As Double
  Dim CurX As Interger
  isRotating = Not isRotating

  If Not isRotating Then Exit sub

  Application.Calculation = xCalculationManual
  Application.ScreenUpdating = True

  inVal = 0.1
  val = 1

  Do
    Sheet5.Shapes(1).Rotation = Sheet5.Shapes(1).Rotation + 50
    DoEvents
  Loop While isRotating

End Sub
```
그림은 PPT에 있으니 다운받아서 돌려보세요

(SigongZoa)
