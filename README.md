# Dolphin-Sync

Dolphin-Sync主要有以下3个应用：

该项目主要由以下三个功能：

一 . 通过vba（Visual Basic for Applications）实现EXCEL数据的zmq传输。【Implementing zmq transmission of EXCEL data through Visual Basic for Applications】

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

代码运行：

![1681811694608](https://user-images.githubusercontent.com/24450492/232741625-bb970134-54ab-4f60-84a4-8522a60fb74c.png)


二. 将EXCEL界面的所有数据转发到python，并进行实时同步。【Forward all data from the EXCEL interface to Python and perform real-time synchronization】
![1681811544540](https://user-images.githubusercontent.com/24450492/232740930-d15e05a0-8f5f-4289-9dca-ac406294eb4a.png)

三. 与EXCEL 插件功能结合，更具有扩展性。【Combined with EXCEL plugin functionality for greater scalability】

1. wind 和 Choice 的EXCEL 插件可以将数据实时更新到EXCEL,其他的Api往往要付费，通过该工具可以将数据同步转发。
2. 将EXCEL作为服务器，不必再进行UI的开发。

![1681815358102](https://user-images.githubusercontent.com/24450492/232756542-eff3caca-04d5-4c2d-b003-2f7f08574348.png)

