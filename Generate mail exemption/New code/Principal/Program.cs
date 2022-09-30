using Principal;

Console.WriteLine("Inicializando...");
DateTime ahora = DateTime.Now;
Boolean deAMentis = true;

Maker make = new Maker(deAMentis);
make.makeAll();

Console.WriteLine(DateTime.Now - ahora);