using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Xml.Serialization;


namespace XMLTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            p.DeserializeObject("C:\\Samples\\XML\\SampleStatementScrubbed.xml");

        
        }

        private void DeserializeObject(string filename)
        {

            XmlSerializer serializer = new XmlSerializer(typeof(statementProduction));

            FileStream fs = new FileStream(filename, FileMode.Open);
            XmlReader reader = new XmlTextReader(fs);

            statementProduction s;

            s = (statementProduction) serializer.Deserialize(reader);

            Console.WriteLine("Institution Name: {0}", s.prologue.institutionName);
            Console.WriteLine("Institution ID: {0}", s.prologue.institutionId);
            Console.WriteLine("Production Date : {0}", s.prologue.productionDate);
            Console.WriteLine("Statement End Date: {0}", s.prologue.statementEndingDate);
            Console.WriteLine("First Name: {0}",s.envelope[0].person.firstName);
            Console.WriteLine("Middle Name : {0}", s.envelope[0].person.middleName);
            Console.WriteLine("Last Name : {0}", s.envelope[0].person.lastName);

            Console.WriteLine("Account Count: {0}", s.epilogue.accountCount);
            Console.WriteLine("Envelope Count : {0}", s.epilogue.envelopeCount);
            Console.ReadLine();
        }
    }
}
