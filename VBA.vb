'2017/10/26
' 语法
数据类型
[布尔Boolean；字节Byte；整数型Integer%；长整数型Long&；单精度浮点型Single！；双精度浮点型Double#；
货币型Currency@；小数型Decimal；字符型String$；日期型Date；对象型Eg:Worksheet]
'声明变量
Dim 变量名 as 数据类型
Private/public/static 变量名 as 数据类型（私有、公有、静态变量）
' 赋值
[let] 变量名=存储数据'数据类型变量赋值
[set] 变量名=存储对象名称'对象类型变量赋值
' example：
1.
Sub 数据变量（）
    Dim IntCount
    IntCount=3000
    Range("A1").value=IntCount'将IntCount中存储的数据写入活动工作表的A1单元格中
End Sub
2.
Sub 对象变量（）
    Dim sht As Worksheet
    Set sht=ActiveSheet
    sht.range("A1").value="我在学习VBA"'在变量sht存储的工作表的A1单元格中输入内容
End Sub

' 数组应用
' 声明
public|Dim 数组名称（a to b) as 数据类型
Dim arr(1 to 100) As Byte   '定义一个Byte类型数组，名称为arr，可以存储100个数据
Dim arr (99) As Byte    '实际上是定义了一个容纳100元素的数组
=Dim arr (0 to 99) As Byte
arr(20)=56  '给数组元素赋值
'声明多维数组
Dim arr(1 to 3,1 to 5) As Integer
Dim arr(2,4) as Integer
=Dim arr(0 to 2,0 to 4) As Integer
'声明动态数组
Dim 数组名称 （）As 数据类型
'example:
Sub Test()
    Dim a As Integer
    a=Application.WorksheetFunction.CountA(Range("A:A"))
    'Application.WorksheetFunction在VBA中使用工作表函数，需要借助Application对象的WorksheetFunction属性来调用
    'Dim arr(1 to a) As String  '错误做法，不能用变量定义数组大小，如果想这样用，必须用动态数组
    Dim arr() AS String
    ReDim arr(1 to a)'ReDim可以重新定义数组大小但是无法更改数组类型
End Sub

'note ：
option Explicit'强制声明
OPTION BASE 1' 模块开始第一句写入"OPTION BASE 1",数组索引号从1开始


'2017/10/27
1.使用Array创建数组
Sub ArrayText()
    Dim arr As Variant  '定义Variant类型变量
    arr = Array(1,2,3,4,5,6,7,8,9,10)
    Msgbox "arr数组的第2个元素为：" & arr(1)
End sub
2.使用Split函数创建数组
' 如果要将一个字符串按指定的分隔符拆开，将各部分结果保存到数组中，可以使用VBA的Split函数。
Sub SplitText()
    Dim arr As Variant  '定义Variant类型变量
    arr = Split("枫叶,空空,小月,老祝",",")'将字符串按都好拆分，并存与数组中，第二参数是采用哪种符号作为分隔符
    Msgbox "arr数组的第2个元素为：" & arr(1)'无论是否在模块写入“OPTION BASE 1”Split函数返回数组索引号都是从0开始
End sub
3.通过单元格直接创建数组
Sub RngArr()
    Dim arr As Variant
    arr = Range("A1:C3").value'将A1:C3中的数据保存到arr中
    Range("E1:G3").value=arr'将数组arr中存储的数据写入E1:G3单元格区域
End Sub

' 数组运算
UBound(数组名称)'return arr max 索引号
LBound(数组名称)'return arr min 索引号
UBound(arr,1)'求数组第一维最大索引号，即x方向；是第一维不是第一列，
UBound(arr,2)'求数组第二维最大索引号，即y方向
UBound-LBound+1'求数组包含的元素个数
' Join函数将一维数组合并为字符串
Sub JoinText()
    Dim arr As Variant,txt As String
    arr = Array(0,1,2,3,4,5,6,7,8,9)
    txt =Join(arr,"@")'用Join函数以@为分隔符，合并数组arr中的元素为一个字符串，将结果保存到变量txt中
    'Join函数第一个参数是要合并的数组名称（只能是一维数组）第二参数是用来分割各元素的分隔符，第二参数默认省略符为空格
    Msgbox txt  '用对话框显示合并数组得到的字符串
    '本例结果为0@1@2@3@4@5@6@7@8@9
End sub
' 将数组中保存的数据写入单元格区域
Range("A1").value=arr(2)    '将数组arr中索引号是2的元素写入活动工作表的A1单元格中
' 可以批量操作
Sub ArrToRng1()
    Dim arr As Variant
    arr = Array(1,2,3,4,5,6,7,8,9,10)
    Range("A1:A9").value=Application.WorksheetFunction.Transpose(arr)'将一维数组写入单元格时，单元格区域必须在同一行，
    '如果要垂直写入一列数据，需要先用工作表的Transpose函数将数组中保存的数据转置为一列。
End sub
' 常量
Const 常量名称 As 数据类型 = 存储在常量中的数据'同样有不同的作用域
' 对象、集合及对象属性和方法
Excel对象层次：工作簿Workbooks-工作表Worksheets-单元格Range
Application.Workbooks("Book1").Worksheet("sheet2").Range("A2")
'Application对象代表Excel程序，是Excel程序最顶岑
' WorkBooks是工作簿集合，代表所有打开的工作簿，Book1是工作簿名称，用来确定要引用工作簿集合中的哪个工作簿
' WorkSheets是工作表集合，代表指定工作簿中的所有工作表；Sheet2是具体要操作的工作表

' 运算符：
+ - * / \ ^ Mod' 算术运算符
' 2.比较运算符
=等于   
<>不等于   
<小于 
>大于 
<=小于等于
Is比较两个对象的引用变量，当对象1和对象2引用相同的对象时返回True，否则返回False
Like比较两个字符串是否匹配
Example:
Range("B2") Like "李*"'B2是否为李开头的任意字符串
VBA中的通配符：
*'代替任意多字符
?'代替任意单个字符
#'代替任意单个数字
[charlist]'代替位于charlist中的任意一个字符，"I" Like "[A-Z]"=True
[!charlist]'代替不在charlist中的任意一个字符，"I" Like "[!H-Z]"=False

' 语句
' 选择语句：
If Range("B2").value >= 60 Then
    Range("C2").value = "及格"
'这里可以插入ElseIf 条件 Then
    '执行语句
Else'最后一个选择用Else
    Range("C2").value = "不及格"
end If
' 用Select Case进行“多选一”
Sub Text()
    Select Case Range("B2").value
        Case Is >=90
            Range("C2").value = "优秀"
        Case Is >=80
            Range("C2").value = "良好"
        Case Is >=60
            Range("C2").value = "及格"
        Case Else
            Range("C2").value = "不及格"
    end Select
end sub
'循环语句
for 循环变量 = 初值 to 终值 Step 步长值
    循环体
    [Exit for]  '提前结束循环
next [循环变量]

For Each 变量 In 集合名称或数组名称
'如果在集合中循环，变量应为对应的对象类型；如果在数组中循环，变量应定义为Variant类型
    语句块1
    [Exit for]
    [语句块2]
Next [变量]
Example:
Sub ShtName()
    Dim sht As Worksheet,i As Integer
    i = 1
    For Each sht In Worksheets
        Range("A" & i) = sht.ShtName
        i=i+1
    Next sht
end sub
' Do while
Do [While]
    <循环体>
    [Exit Do]
    [循环体]
Loop
Do 
    <循环体>
    [Exit Do]
    [循环体]
Loop [While]
' Do Until
Do [Until]
    <循环体>
    [Exit Do]
    [循环体]
Loop
Do 
    <循环体>
    [Exit Do]
    [循环体]
Loop [Until]
' with 语句
Sub FontSet()
    Worksheets("Sheet1").Range("A1").Fount.Name = "仿宋"
    Worksheets("Sheet1").Range("A1").Fount.Size = 12
    Worksheets("Sheet1").Range("A1").Fount.Bold = Ture
    Worksheets("Sheet1").Range("A1").Fount.ColorIndex = 3
end sub
Sub FontSet()
    With Worksheets("Sheet1").Range("A1").Fount
        .Name = "仿宋"    'With与小圆点对应，少了小圆点，则不对该行起作用
        .Size = 3
        .Bold = True
        .ColorIndex = 3
    End with
End Sub

Sub 宏1()
    Range("A1:A8").Select
    Select.Copy '复制选中的区域
    Range("C1").Select  '选择要粘贴区域的起始位置
    ActiveSheet.Paste
End Sub


Sub shtadd()
    Worksheets.Add  '在活动工作表前插入一张新工作表
End sub

note：
Chr(13)'在VBA中的作用相当于按了一次回车
' 使用系统内置函数，VBA代码窗口先输入""VBA."函数列表里会出现待选函数
Range("B" & i)  '用变量表示的单元格位置

