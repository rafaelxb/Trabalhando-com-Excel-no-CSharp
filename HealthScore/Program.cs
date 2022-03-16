using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;



namespace HealthScore
{
    class Program
    {
        static void Main(string[] args)
        {
            //Caminho em que está salvo a planilha
            string path = @"C:\Users\rfxba\Desktop\Desafio CSharp Excel\HealthScore\HealthScore\Model(1).xlsx";
            

            //Le os valores da guia settings
            Excel excel = new Excel(path, 2); //abre o excel na guia correta

            string profiles = excel.ReadValueCell(2, 2);
            string baseHealthScore = excel.ReadValueCell(3, 2);

            excel.Close(); //fecha o excel para não dar problema na hora de salvar


            //Le os valores da guia events
            excel = new Excel(path, 1); 
            
            var listEvents = new List<Events>();

            for (int eachEvent = 2; eachEvent < 7; eachEvent++)
            {
                var events = new Events(eachEvent - 1,
                                        excel.ReadValueCell((eachEvent + 1), 2),
                                        Int32.Parse(excel.ReadValueCell((eachEvent + 1), 3)));
                listEvents.Add(events);
            }

            excel.Close();

            //Grava os dados do cabeçalho na guia output
            excel = new Excel(path, 3);

            foreach (Events events in listEvents)
            {
                excel.SetHeaderValue(events.Id, events.Name);

                excel.Save();
            }
            excel.Close();


            //Grava os dados na guia output
            excel = new Excel(path, 3);

            for (int eachProfile = 1; eachProfile <= Int32.Parse(profiles); eachProfile++)
            {
                Random anyNumber = new Random();

                int healthScore = Int32.Parse(baseHealthScore) + anyNumber.Next(0, 10);

                int pneumonia = 0, breastCancer = 0, hipFracture = 0, parkinsonsDisease = 0, death = 0;

                while (healthScore > 1 && death == 0)
                {
                    int sortEvent = anyNumber.Next(1, 5);
                    healthScore += listEvents[sortEvent - 1].HealthScoreDiscount;

                    if (sortEvent == 1)
                        pneumonia++;
                    else if (sortEvent == 2)
                        breastCancer++;
                    else if (sortEvent == 3)
                        hipFracture++;
                    else if (sortEvent == 4)
                        parkinsonsDisease++;
                    else if (sortEvent == 5)
                        death++;

                }

                    excel.SetValueCell(eachProfile + 2, eachProfile, healthScore, 
                                   pneumonia, breastCancer, hipFracture, 
                                   parkinsonsDisease, death);
            }

            //Salva e fecha a planilha pela última vez
            excel.Save();
            excel.Close();

            Console.WriteLine("Concluído com sucesso");
            
        }

    }

}
