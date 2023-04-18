Attribute VB_Name = "ģ��5"


'ֱ������InitializeZMQ���Է�������
'�����Ҫֹͣ�������ݣ������ `StopZMQSendAsyncWithOnTime` �ӳ���
'����ɷ�����������󣬵��� `CleanUpZMQ` �ӳ����ͷ���Դ��




Option Explicit

' ZeroMQ C API����'
' ����libzmq���е���Ҫ����'
Private Declare PtrSafe Function zmq_ctx_new Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" () As LongPtr
Private Declare PtrSafe Function zmq_ctx_destroy Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr) As Long
Private Declare PtrSafe Function zmq_socket Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr, ByVal stype As Long) As LongPtr
Private Declare PtrSafe Function zmq_bind Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_connect Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_send Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal message As String, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function zmq_close Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr) As Long
Private Declare PtrSafe Function zmq_setsockopt Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal option_name As Long, ByRef option_value As Any, ByVal option_len As Long) As Long

' ZeroMQ ��������'
Private Const ZMQ_PUB As Long = 1
Private Const zmq_NOBLOCK As Long = 1


Dim NextExecution As Double
Dim context As LongPtr
Dim publisher As LongPtr

Sub InitializeZMQ()
    ' ���� ZeroMQ ������'
    context = zmq_ctx_new()
    
    ' ���� ZeroMQ �������׽���'
    publisher = zmq_socket(context, ZMQ_PUB)
    
    ' �󶨵� ZeroMQ ������'
    Dim result As Long
    result = zmq_bind(publisher, "tcp://*:5557")
    Debug.Print "��ʼ��..."
     ZMQSendAsyncWithOnTime
End Sub


Sub ZMQSendAsyncWithOnTime()
    'Debug.Print "��������..."
    Dim Interval As Double
    
    ' ���ü��ʱ�䣬��λΪ��
    Interval = 1 ' 1��

    ' ������һ��ִ�е�ʱ��
    NextExecution = Now + TimeValue("00:00:01") * Interval

    ' ������һ��ִ��
    Application.OnTime NextExecution, "ZMQSendAsyncWithOnTime"

    ' ���÷������ݵķ���
    SendMessage
End Sub

Function ConvertToJson1(sheet_dict As Scripting.Dictionary) As String
    Dim jsonString As String
    Dim sheetName As Variant
    Dim data As Variant
    Dim row As Long
    Dim col As Long

    jsonString = "{"

    For Each sheetName In sheet_dict.Keys
        jsonString = jsonString & """" & sheetName & """: ["
        data = sheet_dict(sheetName)

        For row = LBound(data, 1) To UBound(data, 1)
            jsonString = jsonString & "["
            For col = LBound(data, 2) To UBound(data, 2)
                jsonString = jsonString & """" & Replace(data(row, col), """", """""") & """"
                If col < UBound(data, 2) Then
                    jsonString = jsonString & ","
                End If
            Next col
            jsonString = jsonString & "]"
            If row < UBound(data, 1) Then
                jsonString = jsonString & ","
            End If
        Next row

        jsonString = jsonString & "]"
        If Not sheetName = sheet_dict.Keys(sheet_dict.Count - 1) Then
            jsonString = jsonString & ","
        End If
    Next sheetName

    jsonString = jsonString & "}"
    ConvertToJson1 = jsonString
End Function


Function GetAllSheetData()
    Dim i As Integer
    Dim rowCount As Long
    Dim colCount As Long
    Dim sheet_names As New Collection
    Dim sheet_dict As New Scripting.Dictionary
    
    ' ��ȡ���й����������
    For i = 1 To Worksheets.Count
        sheet_names.Add Worksheets(i).Name
    Next i
    
    ' ���ݹ��������ƻ�ȡ����
    For i = 1 To sheet_names.Count
        Dim current_sheet As Worksheet
        Dim data_range As Range
        Dim current_data As Variant
        
        Set current_sheet = Worksheets(sheet_names(i))
        Set data_range = current_sheet.UsedRange
        
        rowCount = data_range.Rows.Count
        colCount = data_range.Columns.Count
    
        ' ��������������ݣ������ݶ�������
        If rowCount > 1 And colCount > 1 Then
            ' ���������е����ݶ������飨������ʾ��ʽ��
            Dim data() As Variant
            ReDim data(1 To rowCount, 1 To colCount)
            
            Dim k As Long, j As Long
            For k = 1 To rowCount
                For j = 1 To colCount
                    data(k, j) = data_range(k, j).Text
                Next j
            Next k
            
            sheet_dict(current_sheet.Name) = data
        Else
            MsgBox "The worksheet " & current_sheet.Name & " is empty."
        End If
    Next i
    

    Dim json1 As String
    
    json1 = ConvertToJson1(sheet_dict)
    
    GetAllSheetData = json1

End Function





Sub SendMessage()
    Dim result As Long
    Dim message As String


    ' ����Ҫ���͵���Ϣ'
    message = GetAllSheetData()
    Debug.Print "send message: " & Now
    'Debug.Print message
    Debug.Print Len(message)
    
    Dim message_bytes() As Byte
    message_bytes = StrConv(message, vbFromUnicode)


    ' �첽������Ϣ'
    result = zmq_send(publisher, message, UBound(message_bytes) - LBound(message_bytes) + 1, 1)

End Sub

Sub StopZMQSendAsyncWithOnTime()
    Debug.Print "ֹͣ����..."
    ' ֹͣ�ƻ��� OnTime �¼�
    On Error Resume Next
    Application.OnTime NextExecution, "ZMQSendAsyncWithOnTime", , False
    On Error GoTo 0
End Sub

Sub CleanUpZMQ()
    Debug.Print "�ͷ���Դ..."
    ' �ر� ZeroMQ �������׽���'
    Dim result As Long
    result = zmq_close(publisher)
    ' ���� ZeroMQ ������'
    result = zmq_ctx_destroy(context)
End Sub



