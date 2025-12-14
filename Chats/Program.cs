using System;

class Program
{
    static void Main()
    {
        // ryggsack som kan ha 5 saker
        string[] ryggsack = new string[5];
        int val = 0;

        while (val != 4)
        {
            Console.WriteLine("1 Lägg till sak");
            Console.WriteLine("2 Visa saker");
            Console.WriteLine("3 Sök sak");
            Console.WriteLine("4 Avsluta");
            Console.Write("Välj: ");

            if (!int.TryParse(Console.ReadLine(), out val))
            {
                Console.WriteLine("Fel input");
                continue;
            }

            switch (val)
            {
                case 1:
                    // lägga till sak
                    for (int i = 0; i < ryggsack.Length; i++)
                    {
                        if (ryggsack[i] == null)
                        {
                            Console.Write("Skriv sak: ");
                            ryggsack[i] = Console.ReadLine();
                            break;
                        }
                    }
                    break;

                case 2:
                    // visa alla saker
                    for (int i = 0; i < ryggsack.Length; i++)
                    {
                        Console.WriteLine(i + ": " + ryggsack[i]);
                    }
                    break;

                case 3:
                    // sök sak
                    Console.Write("Sök efter: ");
                    string sok = Console.ReadLine();
                    bool hitta = false;

                    for (int i = 0; i < ryggsack.Length; i++)
                    {
                        if (ryggsack[i] == sok)
                        {
                            Console.WriteLine("Hittad på plats " + i);
                            hitta = true;
                        }
                    }

                    if (hitta == false)
                    {
                        Console.WriteLine("Hittade inget");
                    }
                    break;

                case 4:
                    Console.WriteLine("Program slut");
                    break;

                default:
                    Console.WriteLine("Välj mellan 1 och 4");
                    break;
            }
        }
    }
}
