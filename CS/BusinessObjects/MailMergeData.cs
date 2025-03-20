namespace DocumentProcessingWebAPI.BusinessObjects
{
    public class MailMergeData
    {
        readonly static List<Employee> employees = new() {
        new(1, "Nancy", "Davolio", "507 - 20th Ave. E.\nApt. 2A", "Seattle", "98122", "Sales Representative", "(206) 555-9857"),
        new(2, "Andrew", "Fuller", "908 W. Capital Way", "Tacoma", "98401", "Vice President, Sales", "(206) 555-9482"),
        new(3, "Janet", "Leverling", "722 Moss Bay Blvd.", "Kirkland", "98033", "Sales Representative", "(206) 555-3412"),
        new(4, "Margaret", "Peacock", "4110 Old Redmond Rd.", "Redmond", "98052", "Sales Representative", "(206) 555-8122"),
        new(5, "Steven", "Buchanan", "14 Garrett Hill", "London", "SW1 8JR", "Sales Manager", "(71) 555-4848"),
        new(6, "Michael", "Suyama", "Coventry House\nMiner Rd.", "London", "EC2 7JR", "Sales Representative", "(71) 555-7773"),
        new(7, "Robert", "King", "Edgeham Hollow\nWinchester Way", "London", "RG1 9SP", "Sales Representative", "(71) 555-5598"),
        new(8, "Laura", "Callahan", "4726 - 11th Ave. N.E.", "Seattle", "98105", "Inside Sales Coordinator", "(206) 555-1189"),
        new(9, "Anne", "Dodsworth", "7 Houndstooth Rd.", "London", "WG2 7LT", "Sales Representative", "(71) 555-4444"),
    };

        readonly static List<Customer> customers = new(){
        new("ALFKI", "Alfreds Futterkiste", "Maria Anders", "Sales Representative", "Obere Str. 57", "Berlin", "12209"),
        new("ANATR", "Ana Trujillo Emparedados y helados", "Ana Trujillo", "Owner", "Avda. de la Constitución 2222", "México D.F.", "5021"),
        new("ANTON", "Antonio Moreno Taquería", "Antonio Moreno", "Owner", "Mataderos  2312", "México D.F.", "5023"),
        new("AROUT", "Around the Horn", "Thomas Hardy", "Sales Representative", "120 Hanover Sq.", "London", "WA1 1DP"),
        new("BERGS", "Berglunds snabbköp", "Christina Berglund", "Order Administrator", "Berguvsvägen  8", "Luleå", "S-958 22"),
        new("BLAUS", "Blauer See Delikatessen", "Hanna Moos", "Sales Representative", "Forsterstr. 57", "Mannheim", "68306"),
        new("BLONP", "Blondel père et fils", "Frédérique Citeaux", "Marketing Manager", "24, place Kléber", "Strasbourg", "67000"),
        new("BOLID", "Bólido Comidas preparadas", "Martín Sommer", "Owner", "C/ Araquil, 67", "Madrid", "28023"),
        new("BONAP", "Bon app'", "Laurence Lebihan", "Owner", "12, rue des Bouchers", "Marseille", "13008"),
        new("BOTTM", "Bottom-Dollar Markets", "Elizabeth Lincoln", "Accounting Manager", "23 Tsawassen Blvd.", "Tsawwassen", "T2F 8M4"),
        new("BSBEV", "B's Beverages", "Victoria Ashworth", "Sales Representative", "Fauntleroy Circus", "London", "EC2 5NT"),
        new("CACTU", "Cactus Comidas para llevar", "Patricio Simpson", "Sales Agent", "Cerrito 333", "Buenos Aires", "1010"),
        new("CENTC", "Centro comercial Moctezuma", "Francisco Chang", "Marketing Manager", "Sierras de Granada 9993", "México D.F.", "5022"),
        new("CHOPS", "Chop-suey Chinese", "Yang Wang", "Owner", "Hauptstr. 29", "Bern", "3012"),
        new("COMMI", "Comércio Mineiro", "Pedro Afonso", "Sales Associate", "Av. dos Lusíadas, 23", "São Paulo", "05432-043"),
        new("CONSH", "Consolidated Holdings", "Elizabeth Brown", "Sales Representative", "Berkeley Garden\n12  Brewery ", "London", "WX1 6LT"),
        new("DRACD", "Drachenblut Delikatessen", "Sven Ottlieb", "Order Administrator", "Walserweg 21", "Aachen", "52066"),
        new("DUMON", "Du monde entier", "Janine Labrune", "Owner", "67, rue des Cinquante Otages", "Nantes", "44000"),
        new("EASTC", "Eastern Connection", "Ann Devon", "Sales Agent", "35 King George", "London", "WX3 6FW"),
        new("ERNSH", "Ernst Handel", "Roland Mendel", "Sales Manager", "Kirchgasse 6", "Graz", "8010"),
        new("FAMIA", "Familia Arquibaldo", "Aria Cruz", "Marketing Assistant", "Rua Orós, 92", "São Paulo", "05442-030"),
        new("FISSA", "FISSA Fabrica Inter. Salchichas S.A.", "Diego Roel", "Accounting Manager", "C/ Moralzarzal, 86", "Madrid", "28034"),
        new("FOLIG", "Folies gourmandes", "Martine Rancé", "Assistant Sales Agent", "184, chaussée de Tournai", "Lille", "59000"),
        new("FOLKO", "Folk och fä HB", "Maria Larsson", "Owner", "Åkergatan 24", "Bräcke", "S-844 67"),
        new("FRANK", "Frankenversand", "Peter Franken", "Marketing Manager", "Berliner Platz 43", "München", "80805"),
        new("FRANR", "France restauration", "Carine Schmitt", "Marketing Manager", "54, rue Royale", "Nantes", "44000"),
        new("FRANS", "Franchi S.p.A.", "Paolo Accorti", "Sales Representative", "Via Monte Bianco 34", "Torino", "10100"),
        new("FURIB", "Furia Bacalhau e Frutos do Mar", "Lino Rodriguez ", "Sales Manager", "Jardim das rosas n. 32", "Lisboa", "1675"),
        new("GALED", "Galería del gastrónomo", "Eduardo Saavedra", "Marketing Manager", "Rambla de Cataluña, 23", "Barcelona", "8022"),
        new("GODOS", "Godos Cocina Típica", "José Pedro Freyre", "Sales Manager", "C/ Romero, 33", "Sevilla", "41101"),
        new("GOURL", "Gourmet Lanchonetes", "André Fonseca", "Sales Associate", "Av. Brasil, 442", "Campinas", "04876-786"),
        new("GREAL", "Great Lakes Food Market", "Howard Snyder", "Marketing Manager", "2732 Baker Blvd.", "Eugene", "97403"),
        new("GROSR", "GROSELLA-Restaurante", "Manuel Pereira", "Owner", "5ª Ave. Los Palos Grandes", "Caracas", "1081"),
        new("HANAR", "Hanari Carnes", "Mario Pontes", "Accounting Manager", "Rua do Paço, 67", "Rio de Janeiro", "05454-876"),
        new("HILAA", "HILARIÓN-Abastos", "Carlos Hernández", "Sales Representative", "Carrera 22 con Ave. Carlos Soublette #8-35", "San Cristóbal", "5022"),
        new("HUNGC", "Hungry Coyote Import Store", "Yoshi Latimer", "Sales Representative", "City Center Plaza\n516 Main St.", "Elgin", "97827"),
        new("HUNGO", "Hungry Owl All-Night Grocers", "Patricia McKenna", "Sales Associate", "8 Johnstown Road", "Cork", ""),
        new("ISLAT", "Island Trading", "Helen Bennett", "Marketing Manager", "Garden House\nCrowther Way", "Cowes", "PO31 7PJ"),
        new("KOENE", "Königlich Essen", "Philip Cramer", "Sales Associate", "Maubelstr. 90", "Brandenburg", "14776"),
        new("LACOR", "La corne d'abondance", "Daniel Tonini", "Sales Representative", "67, avenue de l'Europe", "Versailles", "78000"),
        new("LAMAI", "La maison d'Asie", "Annette Roulet", "Sales Manager", "1 rue Alsace-Lorraine", "Toulouse", "31000"),
        new("LAUGB", "Laughing Bacchus Wine Cellars", "Yoshi Tannamuri", "Marketing Assistant", "1900 Oak St.", "Vancouver", "V3F 2K1"),
        new("LAZYK", "Lazy K Kountry Store", "John Steel", "Marketing Manager", "12 Orchestra Terrace", "Walla Walla", "99362"),
        new("LEHMS", "Lehmanns Marktstand", "Renate Messner", "Sales Representative", "Magazinweg 7", "Frankfurt a.M. ", "60528"),
        new("LETSS", "Let's Stop N Shop", "Jaime Yorres", "Owner", "87 Polk St.\nSuite 5", "San Francisco", "94117"),
        new("LILAS", "LILA-Supermercado", "Carlos González", "Accounting Manager", "Carrera 52 con Ave. Bolívar #65-98 Llano Largo", "Barquisimeto", "3508"),
        new("LINOD", "LINO-Delicateses", "Felipe Izquierdo", "Owner", "Ave. 5 de Mayo Porlamar", "I. de Margarita", "4980"),
        new("LONEP", "Lonesome Pine Restaurant", "Fran Wilson", "Sales Manager", "89 Chiaroscuro Rd.", "Portland", "97219"),
        new("MAGAA", "Magazzini Alimentari Riuniti", "Giovanni Rovelli", "Marketing Manager", "Via Ludovico il Moro 22", "Bergamo", "24100"),
        new("MAISD", "Maison Dewey", "Catherine Dewey", "Sales Agent", "Rue Joseph-Bens 532", "Bruxelles", "B-1180"),
        new("MEREP", "Mère Paillarde", "Jean Fresnière", "Marketing Assistant", "43 rue St. Laurent", "Montréal", "H1J 1C3"),
        new("MORGK", "Morgenstern Gesundkost", "Alexander Feuer", "Marketing Assistant", "Heerstr. 22", "Leipzig", "4179"),
        new("NORTS", "North/South", "Simon Crowther", "Sales Associate", "South House\n300 Queensbridge", "London", "SW7 1RZ"),
        new("OCEAN", "Océano Atlántico Ltda.", "Yvonne Moncada", "Sales Agent", "Ing. Gustavo Moncada 8585\nPiso 20-A", "Buenos Aires", "1010"),
        new("OLDWO", "Old World Delicatessen", "Rene Phillips", "Sales Representative", "2743 Bering St.", "Anchorage", "99508"),
        new("OTTIK", "Ottilies Käseladen", "Henriette Pfalzheim", "Owner", "Mehrheimerstr. 369", "Köln", "50739"),
        new("PARIS", "Paris spécialités", "Marie Bertrand", "Owner", "265, boulevard Charonne", "Paris", "75012"),
        new("PERIC", "Pericles Comidas clásicas", "Guillermo Fernández", "Sales Representative", "Calle Dr. Jorge Cash 321", "México D.F.", "5033"),
        new("PICCO", "Piccolo und mehr", "Georg Pipps", "Sales Manager", "Geislweg 14", "Salzburg", "5020"),
        new("PRINI", "Princesa Isabel Vinhos", "Isabel de Castro", "Sales Representative", "Estrada da saúde n. 58", "Lisboa", "1756"),
        new("QUEDE", "Que Delícia", "Bernardo Batista", "Accounting Manager", "Rua da Panificadora, 12", "Rio de Janeiro", "02389-673"),
        new("QUEEN", "Queen Cozinha", "Lúcia Carvalho", "Marketing Assistant", "Alameda dos Canàrios, 891", "São Paulo", "05487-020"),
        new("QUICK", "QUICK-Stop", "Horst Kloss", "Accounting Manager", "Taucherstraße 10", "Cunewalde", "1307"),
        new("RANCH", "Rancho grande", "Sergio Gutiérrez", "Sales Representative", "Av. del Libertador 900", "Buenos Aires", "1010"),
        new("RATTC", "Rattlesnake Canyon Grocery", "Paula Wilson", "Assistant Sales Representative", "2817 Milton Dr.", "Albuquerque", "87110"),
        new("REGGC", "Reggiani Caseifici", "Maurizio Moroni", "Sales Associate", "Strada Provinciale 124", "Reggio Emilia", "42100"),
        new("RICAR", "Ricardo Adocicados", "Janete Limeira", "Assistant Sales Agent", "Av. Copacabana, 267", "Rio de Janeiro", "02389-890"),
        new("RICSU", "Richter Supermarkt", "Michael Holz", "Sales Manager", "Grenzacherweg 237", "Genève", "1203"),
        new("ROMEY", "Romero y tomillo", "Alejandra Camino", "Accounting Manager", "Gran Vía, 1", "Madrid", "28001"),
        new("SANTG", "Santé Gourmet", "Jonas Bergulfsen", "Owner", "Erling Skakkes gate 78", "Stavern", "4110"),
        new("SAVEA", "Save-a-lot Markets", "Jose Pavarotti", "Sales Representative", "187 Suffolk Ln.", "Boise", "83720"),
        new("SEVES", "Seven Seas Imports", "Hari Kumar", "Sales Manager", "90 Wadhurst Rd.", "London", "OX15 4NB"),
        new("SIMOB", "Simons bistro", "Jytte Petersen", "Owner", "Vinbæltet 34", "København", "1734"),
        new("SPECD", "Spécialités du monde", "Dominique Perrier", "Marketing Manager", "25, rue Lauriston", "Paris", "75016"),
        new("SPLIR", "Split Rail Beer & Ale", "Art Braunschweiger", "Sales Manager", "P.O. Box 555", "Lander", "82520"),
        new("SUPRD", "Suprêmes délices", "Pascale Cartrain", "Accounting Manager", "Boulevard Tirou, 255", "Charleroi", "B-6000"),
        new("THEBI", "The Big Cheese", "Liz Nixon", "Marketing Manager", "89 Jefferson Way\nSuite 2", "Portland", "97201"),
        new("THECR", "The Cracker Box", "Liu Wong", "Marketing Assistant", "55 Grizzly Peak Rd.", "Butte", "59801"),
        new("TOMSP", "Toms Spezialitäten", "Karin Josephs", "Marketing Manager", "Luisenstr. 48", "Münster", "44087"),
        new("TORTU", "Tortuga Restaurante", "Miguel Angel Paolino", "Owner", "Avda. Azteca 123", "México D.F.", "5033"),
        new("TRADH", "Tradição Hipermercados", "Anabela Domingues", "Sales Representative", "Av. Inês de Castro, 414", "São Paulo", "05634-030"),
        new("TRAIH", "Trail's Head Gourmet Provisioners", "Helvetius Nagy", "Sales Associate", "722 DaVinci Blvd.", "Kirkland", "98034"),
        new("VAFFE", "Vaffeljernet", "Palle Ibsen", "Sales Manager", "Smagsløget 45", "Århus", "8200"),
        new("VICTE", "Victuailles en stock", "Mary Saveley", "Sales Agent", "2, rue du Commerce", "Lyon", "69004"),
        new("VINET", "Vins et alcools Chevalier", "Paul Henriot", "Accounting Manager", "59 rue de l'Abbaye", "Reims", "51100"),
        new("WANDK", "Die Wandernde Kuh", "Rita Müller", "Sales Representative", "Adenauerallee 900", "Stuttgart", "70563"),
        new("WARTH", "Wartian Herkku", "Pirkko Koskitalo", "Accounting Manager", "Torikatu 38", "Oulu", "90110"),
        new("WELLI", "Wellington Importadora", "Paula Parente", "Sales Manager", "Rua do Mercado, 12", "Resende", "08737-363"),
        new("WHITC", "White Clover Markets", "Karl Jablonski", "Owner", "305 - 14th Ave. S.\nSuite 3B", "Seattle", "98128"),
        new("WILMK", "Wilman Kala", "Matti Karttunen", "Owner/Marketing Assistant", "Keskuskatu 45", "Helsinki", "21240"),
        new("WOLZA", "Wolski  Zajazd", "Zbyszek Piestrzeniewicz", "Owner", "ul. Filtrowa 68", "Warszawa", "40909")
    };

        public static List<Employee> Employees => employees;
        public static List<Customer> Customers => customers;
    }

    public class Employee
    {

        public Employee(int employeeId, string firstName, string lastName, string address, string city, string postalCode, string title, string homePhone)
        {
            EmployeeID = employeeId;
            FirstName = firstName;
            LastName = lastName;
            Address = address;
            City = city;
            PostalCode = postalCode;
            Title = title;
            HomePhone = homePhone;
        }

        public int EmployeeID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string FullName => $"{FirstName} {LastName}";
        public string Title { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostalCode { get; set; }
        public string HomePhone { get; set; }

    }

    public class Customer
    {
        public Customer(string customerId, string companyName, string contactName, string contactTitle, string address, string city, string postalCode)
        {
            CustomerID = customerId;
            CompanyName = companyName;
            ContactName = contactName;
            ContactTitle = contactTitle;
            Address = address;
            City = city;
            PostalCode = postalCode;
        }

        public string CustomerID { get; set; }
        public string CompanyName { get; set; }
        public string ContactName { get; set; }
        public string ContactTitle { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostalCode { get; set; }

    }
}
