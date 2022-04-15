﻿using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Net;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace SheetsReader
{
    public class Program
    {
        public static UserCredential Credential;
        public static IList<Sheet> Sheets= new List<Sheet>();
        public static string spreadsheetId = "1ztAJIxFfmCr3QllANpp-UwxPnV4mMceW1tt_UXOAuOk";
        public static void Main(string[] args)
        {
            Credential = GetNewCredential();
            GetAllSheets();
            if (!Directory.Exists("Output"))
            {
                Directory.CreateDirectory("Output");
            }
            DownloadAllSheetsAsCsv();
            
        }

        public static void GetAllSheets()
        {
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = Credential,
                ApplicationName = "Reader"
            });
            
            SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(spreadsheetId);
            Spreadsheet response = request.Execute();
            Sheets = response.Sheets;
        }

        public static int FindGidByName(string findName)
        {
            foreach (Sheet sTmp in Sheets)
            {
                if (sTmp.Properties.Title == findName)
                {
                    return sTmp.Properties.SheetId.Value;
                }
            }
            return -1;
        }

        public static void DownloadSheetAsCsvByName(string findName)
        {
            int gid = FindGidByName(findName);
            if (gid==-1)
            {
                Console.WriteLine("Download failed (Not found "+findName+" on spreadsheet)");
                return;
            }
            Downloader dTmp = new Downloader("https://docs.google.com/spreadsheets/d/"+spreadsheetId+"/export?format=csv&gid=" + gid,  Directory.GetCurrentDirectory()+@"\Output\" + findName+".csv");
            dTmp.StartDownload();
        }

        public static void DownloadAllSheetsAsCsv()
        {
            foreach (Sheet sTmp in Sheets)
            {
                Downloader dTmp = new Downloader("https://docs.google.com/spreadsheets/d/"+spreadsheetId+"/export?format=csv&gid=" + sTmp.Properties.SheetId,  Directory.GetCurrentDirectory()+@"\Output\" + sTmp.Properties.Title+".csv");
                dTmp.StartDownload();
            }
        }
        
        public class Downloader
        {
            private string _url;
            private string _saveLocate;

            public Downloader(string url,string saveLocate)
            {
                _url = url;
                _saveLocate = saveLocate;
            }
            public void StartDownload()
            {
                WebClient client = new WebClient();
                client.DownloadFile(new Uri(_url),_saveLocate);
                Console.WriteLine("Done url: "+_url+" save locate: "+_saveLocate);
            }
        
        }
        public static UserCredential GetNewCredential()
        {
            string[] scope = {SheetsService.Scope.SpreadsheetsReadonly};
            var stream = new FileStream(@"Resources\credentials.json", FileMode.Open, FileAccess.Read);
            return GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, scope, "user", CancellationToken.None, new FileDataStore(@"Resources\", true)).Result;
        }
    }
}