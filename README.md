# Dolphin-Sync
一 . 通过vba（Visual Basic for Applications）实现EXCEL数据的zmq传输；
Excel VBA有多种方式将数据进行广播。
ZeroMQ（也称为PXS MQ、0MQ或zmq）看起来像一个可嵌入的网络库，但其作用类似于一个并发框架。
它为您提供了在各种传输（如进程内、进程间、TCP和多播）上承载原子消息的套接字。
您可以使用pub-sub、任务分发和请求回复等模式将套接字N到N连接起来。
它足够快，可以成为集群各部分。
它的异步I/O模型为您提供了可扩展的多核应用程序，这些应用程序是作为异步消息处理任务构建的。
它有许多语言API，并且在大多数操作系统上运行。

需要提前准备：
1. VBA不直接支持ZeroMQ，需要将C版本的ZeroMQ编译Dll，在VBA中调用。【本代码库给出了WIN10 64位EXCEL 版本】
![1681810506938](https://user-images.githubusercontent.com/24450492/232736279-f90e1ec8-f526-4af5-a249-1fbece6c8816.png)

2. vba页面——工具——引用，增加“Microsoft Scripting Runtime”
![1681811043225](https://user-images.githubusercontent.com/24450492/232738842-18e4bf5c-ad24-4ddc-8e7c-ea664f825d1c.png)

3. 添加VBA-JSON模块
参考：https://github.com/VBA-tools/VBA-JSON




二. 将EXCEL的数据实时转发到python，并进行展示



