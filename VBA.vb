'2017/10/26
' �﷨
��������
[����Boolean���ֽ�Byte��������Integer%����������Long&�������ȸ�����Single����˫���ȸ�����Double#��
������Currency@��С����Decimal���ַ���String$��������Date��������Eg:Worksheet]
'��������
Dim ������ as ��������
Private/public/static ������ as �������ͣ�˽�С����С���̬������
' ��ֵ
[let] ������=�洢����'�������ͱ�����ֵ
[set] ������=�洢��������'�������ͱ�����ֵ
' example��
1.
Sub ���ݱ�������
    Dim IntCount
    IntCount=3000
    Range("A1").value=IntCount'��IntCount�д洢������д���������A1��Ԫ����
End Sub
2.
Sub �����������
    Dim sht As Worksheet
    Set sht=ActiveSheet
    sht.range("A1").value="����ѧϰVBA"'�ڱ���sht�洢�Ĺ������A1��Ԫ������������
End Sub

' ����Ӧ��
' ����
public|Dim �������ƣ�a to b) as ��������
Dim arr(1 to 100) As Byte   '����һ��Byte�������飬����Ϊarr�����Դ洢100������
Dim arr (99) As Byte    'ʵ�����Ƕ�����һ������100Ԫ�ص�����
=Dim arr (0 to 99) As Byte
arr(20)=56  '������Ԫ�ظ�ֵ
'������ά����
Dim arr(1 to 3,1 to 5) As Integer
Dim arr(2,4) as Integer
=Dim arr(0 to 2,0 to 4) As Integer
'������̬����
Dim �������� ����As ��������
'example:
Sub Test()
    Dim a As Integer
    a=Application.WorksheetFunction.CountA(Range("A:A"))
    'Application.WorksheetFunction��VBA��ʹ�ù�����������Ҫ����Application�����WorksheetFunction����������
    'Dim arr(1 to a) As String  '���������������ñ������������С������������ã������ö�̬����
    Dim arr() AS String
    ReDim arr(1 to a)'ReDim�������¶��������С�����޷�������������
End Sub

'note ��
option Explicit'ǿ������
OPTION BASE 1' ģ�鿪ʼ��һ��д��"OPTION BASE 1",���������Ŵ�1��ʼ


'2017/10/27
1.ʹ��Array��������
Sub ArrayText()
    Dim arr As Variant  '����Variant���ͱ���
    arr = Array(1,2,3,4,5,6,7,8,9,10)
    Msgbox "arr����ĵ�2��Ԫ��Ϊ��" & arr(1)
End sub
2.ʹ��Split������������
' ���Ҫ��һ���ַ�����ָ���ķָ����𿪣��������ֽ�����浽�����У�����ʹ��VBA��Split������
Sub SplitText()
    Dim arr As Variant  '����Variant���ͱ���
    arr = Split("��Ҷ,�տ�,С��,��ף",",")'���ַ��������ò�֣������������У��ڶ������ǲ������ַ�����Ϊ�ָ���
    Msgbox "arr����ĵ�2��Ԫ��Ϊ��" & arr(1)'�����Ƿ���ģ��д�롰OPTION BASE 1��Split�����������������Ŷ��Ǵ�0��ʼ
End sub
3.ͨ����Ԫ��ֱ�Ӵ�������
Sub RngArr()
    Dim arr As Variant
    arr = Range("A1:C3").value'��A1:C3�е����ݱ��浽arr��
    Range("E1:G3").value=arr'������arr�д洢������д��E1:G3��Ԫ������
End Sub

' ��������
UBound(��������)'return arr max ������
LBound(��������)'return arr min ������
UBound(arr,1)'�������һά��������ţ���x�����ǵ�һά���ǵ�һ�У�
UBound(arr,2)'������ڶ�ά��������ţ���y����
UBound-LBound+1'�����������Ԫ�ظ���
' Join������һά����ϲ�Ϊ�ַ���
Sub JoinText()
    Dim arr As Variant,txt As String
    arr = Array(0,1,2,3,4,5,6,7,8,9)
    txt =Join(arr,"@")'��Join������@Ϊ�ָ������ϲ�����arr�е�Ԫ��Ϊһ���ַ�������������浽����txt��
    'Join������һ��������Ҫ�ϲ����������ƣ�ֻ����һά���飩�ڶ������������ָ��Ԫ�صķָ������ڶ�����Ĭ��ʡ�Է�Ϊ�ո�
    Msgbox txt  '�öԻ�����ʾ�ϲ�����õ����ַ���
    '�������Ϊ0@1@2@3@4@5@6@7@8@9
End sub
' �������б��������д�뵥Ԫ������
Range("A1").value=arr(2)    '������arr����������2��Ԫ��д���������A1��Ԫ����
' ������������
Sub ArrToRng1()
    Dim arr As Variant
    arr = Array(1,2,3,4,5,6,7,8,9,10)
    Range("A1:A9").value=Application.WorksheetFunction.Transpose(arr)'��һά����д�뵥Ԫ��ʱ����Ԫ�����������ͬһ�У�
    '���Ҫ��ֱд��һ�����ݣ���Ҫ���ù������Transpose�����������б��������ת��Ϊһ�С�
End sub
' ����
Const �������� As �������� = �洢�ڳ����е�����'ͬ���в�ͬ��������
' ���󡢼��ϼ��������Ժͷ���
Excel�����Σ�������Workbooks-������Worksheets-��Ԫ��Range
Application.Workbooks("Book1").Worksheet("sheet2").Range("A2")
'Application�������Excel������Excel������
' WorkBooks�ǹ��������ϣ��������д򿪵Ĺ�������Book1�ǹ��������ƣ�����ȷ��Ҫ���ù����������е��ĸ�������
' WorkSheets�ǹ������ϣ�����ָ���������е����й�����Sheet2�Ǿ���Ҫ�����Ĺ�����

' �������
+ - * / \ ^ Mod' ���������
' 2.�Ƚ������
=����   
<>������   
<С�� 
>���� 
<=С�ڵ���
Is�Ƚ�������������ñ�����������1�Ͷ���2������ͬ�Ķ���ʱ����True�����򷵻�False
Like�Ƚ������ַ����Ƿ�ƥ��
Example:
Range("B2") Like "��*"'B2�Ƿ�Ϊ�ͷ�������ַ���
VBA�е�ͨ�����
*'����������ַ�
?'�������ⵥ���ַ�
#'�������ⵥ������
[charlist]'����λ��charlist�е�����һ���ַ���"I" Like "[A-Z]"=True
[!charlist]'���治��charlist�е�����һ���ַ���"I" Like "[!H-Z]"=False

' ���
' ѡ����䣺
If Range("B2").value >= 60 Then
    Range("C2").value = "����"
'������Բ���ElseIf ���� Then
    'ִ�����
Else'���һ��ѡ����Else
    Range("C2").value = "������"
end If
' ��Select Case���С���ѡһ��
Sub Text()
    Select Case Range("B2").value
        Case Is >=90
            Range("C2").value = "����"
        Case Is >=80
            Range("C2").value = "����"
        Case Is >=60
            Range("C2").value = "����"
        Case Else
            Range("C2").value = "������"
    end Select
end sub
'ѭ�����
for ѭ������ = ��ֵ to ��ֵ Step ����ֵ
    ѭ����
    [Exit for]  '��ǰ����ѭ��
next [ѭ������]

For Each ���� In �������ƻ���������
'����ڼ�����ѭ��������ӦΪ��Ӧ�Ķ������ͣ������������ѭ��������Ӧ����ΪVariant����
    ����1
    [Exit for]
    [����2]
Next [����]
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
    <ѭ����>
    [Exit Do]
    [ѭ����]
Loop
Do 
    <ѭ����>
    [Exit Do]
    [ѭ����]
Loop [While]
' Do Until
Do [Until]
    <ѭ����>
    [Exit Do]
    [ѭ����]
Loop
Do 
    <ѭ����>
    [Exit Do]
    [ѭ����]
Loop [Until]
' with ���
Sub FontSet()
    Worksheets("Sheet1").Range("A1").Fount.Name = "����"
    Worksheets("Sheet1").Range("A1").Fount.Size = 12
    Worksheets("Sheet1").Range("A1").Fount.Bold = Ture
    Worksheets("Sheet1").Range("A1").Fount.ColorIndex = 3
end sub
Sub FontSet()
    With Worksheets("Sheet1").Range("A1").Fount
        .Name = "����"    'With��СԲ���Ӧ������СԲ�㣬�򲻶Ը���������
        .Size = 3
        .Bold = True
        .ColorIndex = 3
    End with
End Sub

Sub ��1()
    Range("A1:A8").Select
    Select.Copy '����ѡ�е�����
    Range("C1").Select  'ѡ��Ҫճ���������ʼλ��
    ActiveSheet.Paste
End Sub


Sub shtadd()
    Worksheets.Add  '�ڻ������ǰ����һ���¹�����
End sub

note��
Chr(13)'��VBA�е������൱�ڰ���һ�λس�
' ʹ��ϵͳ���ú�����VBA���봰��������""VBA."�����б������ִ�ѡ����
Range("B" & i)  '�ñ�����ʾ�ĵ�Ԫ��λ��

