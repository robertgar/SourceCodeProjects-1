using Principal;
Console.WriteLine("Inicializando...");
DateTime ahora = DateTime.Now;

Maker make = new Maker();
make.makeAll();

Console.WriteLine(DateTime.Now - ahora);