Attribute VB_Name = "模块1"
Option Explicit

' ZeroMQ C API声明'
' 声明libzmq库中的主要函数'
Private Declare PtrSafe Function zmq_ctx_new Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" () As LongPtr
Private Declare PtrSafe Function zmq_ctx_destroy Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr) As Long
Private Declare PtrSafe Function zmq_socket Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal context As LongPtr, ByVal stype As Long) As LongPtr
Private Declare PtrSafe Function zmq_bind Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_connect Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal endpoint As String) As Long
Private Declare PtrSafe Function zmq_send Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal message As String, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function zmq_recv Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByRef buffer As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function zmq_close Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr) As Long
Private Declare PtrSafe Function zmq_setsockopt Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal option_name As Long, ByRef option_value As Any, ByVal option_len As Long) As Long
Private Declare PtrSafe Function zmq_poll Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByRef items As ZMQPollItem, ByVal nitems As Long, ByVal timeout As Long) As Long
Private Declare PtrSafe Function zmq_getsockopt Lib "C:\Users\llx\Desktop\zmq\libzmq.dll" (ByVal socket As LongPtr, ByVal option_name As Long, ByRef option_value As Any, ByRef option_len As Long) As Long


' ZeroMQ 常量定义'
Private Const ZMQ_SUB As Long = 2
Private Const ZMQ_POLLIN As Long = 1

' ZMQPollItem 类型定义'
Private Type ZMQPollItem
    socket As LongPtr
    fd As LongPtr
    events As Long
    revents As Long
End Type

' 回调函数'
Private Sub OnMessageReceived(ByVal message As String)
    Debug.Print "Received message: " & message
End Sub

' 异步接收数据'
Private Sub ZMQReceiveAsync()
    ' 定义变量'
    Dim context As LongPtr
    Dim subscriber As LongPtr
    Dim result As Long
    Dim message As String
    Dim pollItems(0 To 0) As ZMQPollItem
    Dim timeout As Long
    Dim callback As String
    
    ' 创建 ZeroMQ 上下文'
    context = zmq_ctx_new()
    
    ' 创建 ZeroMQ 订阅者套接字'
    subscriber = zmq_socket(context, ZMQ_SUB) '
    ' 连接到 ZeroMQ 服务器'
    result = zmq_connect(subscriber, "tcp://localhost:5555")
    ' 设置订阅规则'
    Dim topic As String
    topic = "topic"
    result = zmq_setsockopt(subscriber, 6, ByVal topic, Len(topic)) ' 订阅 "topic" 主题
    
    ' 设置超时时间为 50 毫秒'
    timeout = 50
    
    ' 设置 pollItems 数组'
    pollItems(0).socket = subscriber
    pollItems(0).events = ZMQ_POLLIN
    
    ' 设置回调函数名称'
    callback = "OnMessageReceived"
    Debug.Print "waiting... "
    

    Do While True
        result = zmq_poll(pollItems(0), UBound(pollItems) + 1, timeout)

        If result > 0 Then
            If ZMQ_POLLIN Then
                message = Space$(256)
                result = zmq_recv(subscriber, ByVal message, Len(message), 0)
                Range("B2").Value = "Received message: " & Left$(message, result)
                
            'message = StrConv(Left$(message, result), vbUnicode)
                Debug.Print "Received message: " & Left$(message, result)
            End If
        End If
        
        ' 允许程序进行其他操作
        DoEvents
    Loop
    


    
    ' 关闭 ZeroMQ 订阅者套接字'
    result = zmq_close(subscriber)
    
    ' 销毁 ZeroMQ 上下文'
    result = zmq_ctx_destroy(context)
End Sub

