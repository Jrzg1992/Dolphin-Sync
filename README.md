The Dolphin-Sync project is expected to provide some inspiration for everyone in "communicating with other applications in Excel". This project has not undergone stress testing and although it can run normally, there may be many bugs. Please refer to it as appropriate.

This project mainly consists of the following three functions:

I Realize zmq transmission of EXCEL data through Visual Basic for Applications.

Excel VBA has multiple ways to broadcast data.


ZeroMQ (also known as PXS MQ, 0MQ, or zmq) looks like an embeddable network library, but its function is similar to a concurrency framework.



It provides you with sockets that host atomic messages on various transports such as intra process, inter process, TCP, and multicast.



You can use patterns such as pub sub, task distribution, and request reply to connect sockets N to N.



It is fast enough to become part of the cluster.



Its asynchronous I/O model provides you with scalable multi-core applications built as asynchronous message processing tasks.



It has many language APIs and runs on most operating systems.



Need to prepare in advance:

1. VBA does not directly support ZeroMQ. You need to compile ZeroMQ version C into Dll and call it in VBA. This code library provides the WIN10 64 bit EXCEL version



! [1681810506938]( https://user-images.githubusercontent.com/24450492/232736279-f90e1ec8-f526-4af5-a249-1fbece6c8816.png )



2. vba page - Tools - Reference, adding 'Microsoft Scripting Runtime'



! [1681811043225]( https://user-images.githubusercontent.com/24450492/232738842-18e4bf5c-ad24-4ddc-8e7c-ea664f825d1c.png )



3. Add VBA-JSON module

reference resources: https://github.com/VBA-tools/VBA-JSON



Code Run:



! [1681811694608]( https://user-images.githubusercontent.com/24450492/232741625-bb970134-54ab-4f60-84a4-8522a60fb74c.png )




II Forward all data from the EXCEL interface to Python and perform real-time synchronization.



! [1681811544540]( https://user-images.githubusercontent.com/24450492/232740930-d15e05a0-8f5f-4289-9dca-ac406294eb4a.png )



III Combined with the EXCEL plugin function, it has more scalability.



1. The Excel plugins of Wind and Choice can update data in real-time to Excel, while other Api tools often require payment and can synchronize and forward data through this tool.

2. Using EXCEL as a server eliminates the need for UI development.



! [1681815358102]( https://user-images.githubusercontent.com/24450492/232756542-eff3caca-04d5-4c2d-b003-2f7f08574348.png )
