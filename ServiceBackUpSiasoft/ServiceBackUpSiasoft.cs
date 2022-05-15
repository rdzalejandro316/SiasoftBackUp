using Azure.Storage.Blobs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace ServiceBackUpSiasoft
{
    public partial class ServiceBackUpSiasoft : ServiceBase
    {

        private System.Timers.Timer Timer = null;

        /// intervalos de tiempo para la ejecucion de las funciones        
        bool OneTime = Convert.ToBoolean(ConfigurationManager.AppSettings["OneTime"]);
        bool IsMinutes = Convert.ToBoolean(ConfigurationManager.AppSettings["IsMinutes"]);
        int IntervalInMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalInMinutes"]);
        bool IsShedule = Convert.ToBoolean(ConfigurationManager.AppSettings["IsShedule"]);
        int IntervalShedule = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalShedule"]);

        /// ruta de backups 
        string PathBackUp = ConfigurationManager.AppSettings["PathBackUp"];        
        string PathMoveZip = ConfigurationManager.AppSettings["PathMoveZip"];

        /// cuenta de azure        
        string AzureCnBlob = ConfigurationManager.AppSettings["AzureCnBlob"];
        string ContainerAzureBlob = ConfigurationManager.AppSettings["ContainerAzureBlob"];

       

        public ServiceBackUpSiasoft()
        {
            InitializeComponent();
        }

        protected override async void OnStart(string[] args)
        {
            Message("INICIO EL SERVICIO DE BACKUP");

            if (OneTime)
            {
                await ExecuteTask();
            }
            else if (IsMinutes)
            {
                Timer = new System.Timers.Timer();
                Timer.Interval = TimeSpan.FromMinutes(IntervalInMinutes).TotalMilliseconds;
                Timer.Elapsed += Timer_Elapsed;
                Timer.Start();
            }
            else if (IsShedule)
            {
                Timer = new System.Timers.Timer();
                // revisa cada 40 minutos para que entre en cada hora del dia
                Timer.Interval = TimeSpan.FromMinutes(60).TotalMilliseconds;
                Timer.Elapsed += Timer_Elapsed;
                Timer.Start();

            }
        }

        private async void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (IsShedule)
                {
                    DateTime date = DateTime.Now;
                    int hora = Convert.ToInt32(date.ToString("HH"));
                    if (hora != IntervalShedule) return;
                }

                await ExecuteTask();
            }
            catch (Exception w)
            {
                Message("error al ejecutar Timer_Elapsed():" + w, EventLogEntryType.Error);
            }
        }

        public async Task ExecuteTask()
        {
            try
            {
                Message("INICIO LA COMPRESION ZIP");
                var ZipBackUp = Task.Factory.StartNew(() => IniZip());
                await ZipBackUp;

                if (ZipBackUp.IsCompleted)
                {
                    Message("INICIO LA SUBIDA A AZURE BLOB STORAGE");
                    var blob =  UpLoadZipAzureBlob();
                    await blob;
                    
                    if (blob.IsCompleted)
                    {
                        Message("INICIO EL DESPLAZAMIENTO DE LOS ARCHIVOS ZIP");
                        var MoveZip = Task.Factory.StartNew(() => MoveZipPrevious());
                        await MoveZip;

                        if (MoveZip.IsCompleted)
                        {
                            Message("INICIO LA ELIMINACION DEL .BAK");
                            var DeleteBak = Task.Factory.StartNew(() => IniDeleteFileResult());
                            await DeleteBak;
                        }                      
                    }

                }               
                Message("FINALIZO PROCESO DEL SERVICIO");
            }
            catch (Exception w)
            {
                Message("error al ejecutar Timer_Elapsed():" + w, EventLogEntryType.Error);
            }
        }

        #region LOGIC EVENTS

        public void IniZip()
        {
            try
            {

                DirectoryInfo d = new DirectoryInfo(PathBackUp);
                FileInfo[] FilesBak = d.GetFiles("*.bak");
                foreach (FileInfo file in FilesBak)
                {
                    string pathFile = file.FullName;
                    string fileDirectory = Path.GetDirectoryName(pathFile);
                    string filSinExten = Path.GetFileNameWithoutExtension(pathFile);
                    string zipPath = $@"{fileDirectory}\{filSinExten}.zip";

                    using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                        zip.CreateEntryFromFile(pathFile, System.IO.Path.GetFileName(pathFile));
                }
            }
            catch (Exception w)
            {
                EventLog.WriteEntry("error al comprimir:" + w, EventLogEntryType.Error);
            }
        }

        public async Task UpLoadZipAzureBlob()
        {
            try
            {
                BlobServiceClient blobServiceClient = new BlobServiceClient(AzureCnBlob);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(ContainerAzureBlob);

                DirectoryInfo d = new DirectoryInfo(PathBackUp);
                FileInfo[] Files = d.GetFiles("*.zip");
                foreach (FileInfo file in Files)
                {
                    string pathFile = file.FullName;
                    string name = Path.GetFileName(pathFile);
                    string paths = Path.GetFullPath(pathFile);

                    BlobClient blobClient = containerClient.GetBlobClient(name);
                    using (var fileStream = System.IO.File.OpenRead(paths))
                    {
                        await blobClient.UploadAsync(fileStream, true);
                    }
                }

            }
            catch (Exception w)
            {
                Message("error al subir a azure blob storage:" + w, EventLogEntryType.Error);
            }
        }

        public void MoveZipPrevious()
        {
            try
            {
                List<FolderZip> MoveZip = new List<FolderZip>();

                DirectoryInfo d = new DirectoryInfo(PathBackUp);
                FileInfo[] FilesBak = d.GetFiles("*.zip");
                foreach (FileInfo file in FilesBak)
                {
                    FolderZip folderZip = new FolderZip();
                    folderZip.from = file.FullName;
                    folderZip.to = $@"{PathMoveZip}\{Path.GetFileName(file.FullName)}";
                    MoveZip.Add(folderZip);
                }

                foreach (FolderZip item in MoveZip)
                    System.IO.File.Move(item.from, item.to);
            }
            catch (Exception w)
            {
                EventLog.WriteEntry("error al comprimir:" + w, EventLogEntryType.Error);
            }
        }

        public void IniDeleteFileResult()
        {
            try
            {
                DirectoryInfo d = new DirectoryInfo(PathBackUp);
                FileInfo[] FilesBak = d.GetFiles("*.bak");

                List<string> fileDelete = new List<string>();
                foreach (FileInfo file in FilesBak)
                    fileDelete.Add(file.FullName);

                if (fileDelete.Count > 0)
                {
                    foreach (var file in fileDelete)
                        File.Delete(file);
                }
            }
            catch (Exception w)
            {
                EventLog.WriteEntry("error al comprimir:" + w, EventLogEntryType.Error);
            }
        }

        #endregion


        protected override void OnStop()
        {
            Timer.Dispose();
            Message("FINALIZO EL SERVICIO DE ServiceBackUpSiasoft");
        }

        public void Message(string message, EventLogEntryType tipo = EventLogEntryType.Information)
        {
            EventLog.WriteEntry(message, tipo);
        }

        public class FolderZip
        {
            public string from { get; set; }
            public string to { get; set; }
        }
    }
}
