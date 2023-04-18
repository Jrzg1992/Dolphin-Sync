Attribute VB_Name = "模块5"


'直接运行InitializeZMQ可以发放数据
'如果需要停止发送数据，请调用 `StopZMQSendAsyncWithOnTime` 子程序。
'在完成发送数据任务后，调用 `CleanUpZMQ` 子程序释放资源。




Option Explicit

' ZeroMQ C API声明'
' 声明libzmq库中的主要函数'
Private Declare PtrSafe Function zmq_ctx_new Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" () As LongPtr
Private Declare PtrSafe Function zmq_ctx_destroy Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr) As Long
Private Declare PtrSafe Function zmq_socket Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr, ByVal stype As Long) As LongPtr
Private Declare PtrSafe Function zmq_bind Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_connect Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_send Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal message As String, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function zmq_close Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr) As Long
Private Declare PtrSafe Function zmq_setsockopt Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal option_name As Long, ByRef option_value As Any, ByVal option_len As Long) As Long

' ZeroMQ 常量定义'
Private Const ZMQ_PUB As Long = 1
Private Const zmq_NOBLOCK As Long = 1


Dim NextExecution As Double
Dim context As LongPtr
Dim publisher As LongPtr

Sub InitializeZMQ()
    ' 创建 ZeroMQ 上下文'
    context = zmq_ctx_new()
    
    ' 创建 ZeroMQ 发布者套接字'
    publisher = zmq_socket(context, ZMQ_PUB)
    
    ' 绑定到 ZeroMQ 服务器'
    Dim result As Long
    result = zmq_bind(publisher, "tcp://*:5557")
    Debug.Print "初始化..."
     ZMQSendAsyncWithOnTime
End Sub


Sub ZMQSendAsyncWithOnTime()
    'Debug.Print "发送数据..."
    Dim Interval As Double
    
    ' 设置间隔时间，单位为秒
    Interval = 1 ' 1秒

    ' 计算下一次执行的时间
    NextExecution = Now + TimeValue("00:00:01") * Interval

    ' 调度下一次执行
    Application.OnTime NextExecution, "ZMQSendAsyncWithOnTime"

    ' 调用发送数据的方法
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
    
    ' 获取所有工作表的名称
    For i = 1 To Worksheets.Count
        sheet_names.Add Worksheets(i).Name
    Next i
    
    ' 根据工作表名称获取数据
    For i = 1 To sheet_names.Count
        Dim current_sheet As Worksheet
        Dim data_range As Range
        Dim current_data As Variant
        
        Set current_sheet = Worksheets(sheet_names(i))
        Set data_range = current_sheet.UsedRange
        
        rowCount = data_range.Rows.Count
        colCount = data_range.Columns.Count
    
        ' 如果工作表有数据，则将数据读入数组
        If rowCount > 1 And colCount > 1 Then
            ' 将工作表中的数据读入数组（包括显示格式）
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


    ' 构造要发送的消息'
    message = GetAllSheetData()
    Debug.Print "send message: " & Now
    'Debug.Print message
    Debug.Print Len(message)
    
    Dim message_bytes() As Byte
    message_bytes = StrConv(message, vbFromUnicode)


    ' 异步发送消息'
    result = zmq_send(publisher, message, UBound(message_bytes) - LBound(message_bytes) + 1, 1)

End Sub

Sub StopZMQSendAsyncWithOnTime()
    Debug.Print "停止数据..."
    ' 停止计划的 OnTime 事件
    On Error Resume Next
    Application.OnTime NextExecution, "ZMQSendAsyncWithOnTime", , False
    On Error GoTo 0
End Sub

Sub CleanUpZMQ()
    Debug.Print "释放资源..."
    ' 关闭 ZeroMQ 发布者套接字'
    Dim result As Long
    result = zmq_close(publisher)
    ' 销毁 ZeroMQ 上下文'
    result = zmq_ctx_destroy(context)
End Sub



