VegaCOMM Notes

Installing a module:
To install a module, simple click on the Modules menu option, select Module Administration. One the the module window is open, select Module you wish to isntall, right click and select isntall.

For PSC users:

Due to the influx of worms/trojans etc that appear on PSC I understand the worries of using precompiled objects. so I have included just the source for everything I used. Take these steps to get it ready to work.

1) Compile and register the UserModule.dll. The code is in the Server/UserModule directory

2) To use the example modules. goto the Message Admim Module directory. Compile the code in the client directory to the vegacomm\client\Modules directory. And compile the code in the server directory to the Vegacomm\server\modules. See above to see how to install them in VegaCOMM.

If you wish to learn how to make modules, check out the ModuleSDK in  the Server/Modules directory. It is not complete yet, buy you can get a good idea form the code on the server on how to write a module.

