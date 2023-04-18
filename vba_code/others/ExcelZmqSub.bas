Attribute VB_Name = "ģ��1"
Option Explicit

' ZeroMQ C API����'
' ����libzmq���е���Ҫ����'
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


' ZeroMQ ��������'
Private Const ZMQ_SUB As Long = 2
Private Const ZMQ_POLLIN As Long = 1

' ZMQPollItem ���Ͷ���'
Private Type ZMQPollItem
    socket As LongPtr
    fd As LongPtr
    events As Long
    revents As Long
End Type

' �ص�����'
Private Sub OnMessageReceived(ByVal message As String)
    Debug.Print "Received message: " & message
End Sub

' �첽��������'
Private Sub ZMQReceiveAsync()
    ' �������'
    Dim context As LongPtr
    Dim subscriber As LongPtr
    Dim result As Long
    Dim message As String
    Dim pollItems(0 To 0) As ZMQPollItem
    Dim timeout As Long
    Dim callback As String
    
    ' ���� ZeroMQ ������'
    context = zmq_ctx_new()
    
    ' ���� ZeroMQ �������׽���'
    subscriber = zmq_socket(context, ZMQ_SUB) '
    ' ���ӵ� ZeroMQ ������'
    result = zmq_connect(subscriber, "tcp://localhost:5555")
    ' ���ö��Ĺ���'
    Dim topic As String
    topic = "topic"
    result = zmq_setsockopt(subscriber, 6, ByVal topic, Len(topic)) ' ���� "topic" ����
    
    ' ���ó�ʱʱ��Ϊ 50 ����'
    timeout = 50
    
    ' ���� pollItems ����'
    pollItems(0).socket = subscriber
    pollItems(0).events = ZMQ_POLLIN
    
    ' ���ûص���������'
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
        
        ' ������������������
        DoEvents
    Loop
    


    
    ' �ر� ZeroMQ �������׽���'
    result = zmq_close(subscriber)
    
    ' ���� ZeroMQ ������'
    result = zmq_ctx_destroy(context)
End Sub

