using Principal;

Boolean deAMentis = true;
DateTime justNow = DateTime.Now;

Maker make = new Maker(deAMentis);

make.Begin(justNow);
make.makeAll();
make.End(justNow);