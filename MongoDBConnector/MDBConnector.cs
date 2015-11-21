using System;
using MongoDB.Bson;
using MongoDB.Driver;

namespace MongoDBConnector
{
    using System.Collections.Generic;
    using global::MongoDB.Driver.Builders;

    internal sealed class MDBConnector
    {
        private MongoClient client;
        private MongoServer _server;
        private readonly string _database;
        private MongoServerSettings _settings;
        
        public void ConnectMongoDB()
        {
            _server = new MongoServer(settings);
            _database = databaseName;
            _server.Connect();
            client.GetServer();
        }

        private MongoServerSettings SetMongoServerSettings()
        {
            return new MongoServerSettings()
            {
                
            }
        }
    }
}
