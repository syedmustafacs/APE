using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using MongoDB.Driver.GridFS;
using MongoDB.Driver.Linq; 



namespace APE_ManagementSystem
{
    //mongodb://127.0.0.1:27017
    // 1,4,6,8,10,11,16...1,4,6,10,11,16
    public class Symbol  
{  
public string Name { get; set; }  
public ObjectId ID { get; set; }  
}

    
 class  DbConnection

    {
        public static MongoDatabase conn(){
            String connectionString = "mongodb://127.0.0.1:27017";
            MongoClient client = new MongoClient(connectionString);
            MongoServer server = client.GetServer();
            MongoDatabase database = server.GetDatabase("UserManagement");
            return database;
        }
        public void connection() {
            String connectionString = "mongodb://127.0.0.1:27017";
            MongoClient client = new MongoClient(connectionString);
            MongoServer server = client.GetServer();
            MongoDatabase database = server.GetDatabase("UserManagement");
            MongoCollection symbolcollection = database.GetCollection<Symbol>("Profile");  
            Symbol symbol = new Symbol ();  
            
            symbol.Name = "Star";  
                symbolcollection.Insert(symbol); 
        }
    }
}
