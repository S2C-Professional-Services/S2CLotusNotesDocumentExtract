using Domino;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace S2CLotusNotesExtractTool
{
    public class LotusNotesIntegrator
    {
        private string _lotusCientPassword = null;
        private string _lotusnotesserverName = null;
        private bool _IsfetchServerData = false;

        private static readonly string BaseDirectoryPath = @"E:\Documents\";
        public LotusNotesIntegrator(string Password, string ServerName, bool isFetchServerData)
        {
            _lotusCientPassword = Password;
            _lotusnotesserverName = ServerName;
            _IsfetchServerData = isFetchServerData;
        }
        public string ClientPassword
        {
            get
            {
                return _lotusCientPassword;
            }
        }
        public string LotusNotesServer
        {
            get
            {
                return _lotusnotesserverName;
            }
        }
        public bool FetchServerData
        {
            get
            {
                return _IsfetchServerData;
            }
        }

        public void ExtractGraphicsCafeDocuments()
        {
            var specificCategories = ReadSpecificCategories();
            Console.WriteLine("Initialize the Extract with Cred: Server: " + LotusNotesServer + " Password: " + ClientPassword);
            try
            {
                NotesSession s = new Domino.NotesSessionClass();
                s.Initialize(ClientPassword);
                NotesDbDirectory d = s.GetDbDirectory(LotusNotesServer);
                NotesDatabase db = d.GetFirstDatabase(DB_TYPES.NOTES_DATABASE);

                bool isAlive = true;
                while (isAlive)
                {
                    Console.WriteLine("Reading file: " + db.FileName);
                    if (db != null && db.FileName.ToLower() == "Graphics Cafe.nsf".ToLower())
                    {
                        Console.WriteLine("Reading file: " + db.FileName);
                        isAlive = false;
                        try
                        {
                            if (!db.IsOpen)
                                db.Open();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Unable to open database: " + db.FileName);
                            Console.WriteLine("Exception: " + ex);
                            db = d.GetNextDatabase();
                            continue;
                        }
                        NotesDocumentCollection docs = db.AllDocuments;

                        if (docs != null)
                        {
                            Console.WriteLine("Documents Count: " + db.AllDocuments.Count);
                            string rootFolderName = Path.GetFileNameWithoutExtension(db.FileName);
                            string baseDirectoryPath = BaseDirectoryPath + rootFolderName + "\\";

                            for (int rowCount = 1; rowCount <= docs.Count; rowCount++)
                            {
                                NotesDocument document = docs.GetNthDocument(rowCount);
                                if (document.HasEmbedded && document.HasItem("$File"))
                                {
                                    object[] AllDocItems = (object[])document.Items;
                                    string category = string.Empty;
                                    List<string> categories = new List<string>();
                                    string subCategory = string.Empty;
                                    List<string> subCategories = new List<string>();
                                    Console.WriteLine("RowIndex: " + rowCount);
                                    string documentDescription = string.Empty;
                                    string documentDate = string.Empty;
                                    string documentNotes = string.Empty;
                                    string documentAuthor = string.Empty;
                                    var notesItems = new List<NotesItem>();
                                    foreach (object CurItem in AllDocItems)
                                    {
                                        notesItems.Add((NotesItem)CurItem);
                                    }

                                    var categoryItem = notesItems.FirstOrDefault(x => x.Name == "Category");
                                    if (categoryItem != null && !string.IsNullOrEmpty(categoryItem.Text))
                                    {
                                        category = categoryItem.Text;
                                    }
                                    var subcategoryItem = notesItems.FirstOrDefault(x => x.Name == "SubCategory");

                                    if (subcategoryItem != null && !string.IsNullOrEmpty(subcategoryItem.Text))
                                    {
                                        subCategory = subcategoryItem.Text;
                                    }

                                    var descriptionItem = notesItems.FirstOrDefault(x => x.Name == "Description");
                                    if (descriptionItem != null && !string.IsNullOrEmpty(descriptionItem.Text))
                                    {
                                        documentDescription = descriptionItem.Text;
                                    }

                                    var notesItem = notesItems.FirstOrDefault(x => x.Name == "Notes");
                                    if (notesItem != null && !string.IsNullOrEmpty(notesItem.Text))
                                    {
                                        documentNotes = notesItem.Text;
                                    }

                                    var dateItem = notesItems.FirstOrDefault(x => x.Name == "tmpDate");
                                    if (dateItem != null && !string.IsNullOrEmpty(dateItem.Text))
                                    {
                                        documentDate = dateItem.Text;
                                    }

                                    var authorItem = notesItems.FirstOrDefault(x => x.Name == "tmpAuthor");
                                    if (authorItem != null && !string.IsNullOrEmpty(authorItem.Text))
                                    {
                                        documentAuthor = authorItem.Text;
                                    }

                                    if (!string.IsNullOrEmpty(category))
                                    {
                                        categories = category.Split(';').ToList();
                                    }
                                    if (!string.IsNullOrEmpty(subCategory))
                                    {
                                        subCategories = subCategory.Split(';').ToList();
                                    }

                                    if (specificCategories.Count == 0 || (specificCategories.Count > 0 && (categories.Any(x=> specificCategories.Contains(x)) || subCategories.Any(x=> specificCategories.Contains(x)))))
                                    {
                                        List<NotesItem> documentItems = notesItems.Where(x => x.type == IT_TYPE.ATTACHMENT).ToList();
                                        if (documentItems != null && documentItems.Count > 0)
                                        {
                                            foreach (var nItem in documentItems)
                                            {
                                                if (IT_TYPE.ATTACHMENT == nItem.type)
                                                {
                                                    var pAttachment = ((object[])nItem.Values)[0].ToString();
                                                    Console.WriteLine("Description: " + documentDescription);
                                                    Console.WriteLine("Date: " + documentDate);
                                                    Console.WriteLine("Notes: " + documentNotes);
                                                    Console.WriteLine("Author: " + documentAuthor);
                                                    foreach (var cat in categories)
                                                    {
                                                        Console.WriteLine("Category: " + cat);
                                                        string destPath = baseDirectoryPath;
                                                        if (!string.IsNullOrEmpty(cat))
                                                        {
                                                            if (!Directory.Exists(baseDirectoryPath + cat))
                                                            {
                                                                Directory.CreateDirectory(baseDirectoryPath + cat);
                                                            }
                                                            destPath = destPath + cat + "\\";
                                                        }

                                                        foreach (var subcat in subCategories)
                                                        {
                                                            Console.WriteLine("SubCategory: " + subcat);

                                                            if (!string.IsNullOrEmpty(subcat))
                                                            {
                                                                if (!Directory.Exists(baseDirectoryPath + cat + "\\" + subcat))
                                                                {
                                                                    Directory.CreateDirectory(baseDirectoryPath + cat + "\\" + subcat);
                                                                }
                                                                destPath = destPath + subcat + "\\";
                                                            }
                                                        }
                                                        Console.WriteLine("Final Destination Path: " + destPath + pAttachment);
                                                        if (!File.Exists(destPath + pAttachment))
                                                        {
                                                            try
                                                            {

                                                                document.GetAttachment(pAttachment).ExtractFile(destPath + pAttachment);

                                                            }
                                                            catch (Exception exe)
                                                            {
                                                                LogDetails("File not extracted: " + destPath + pAttachment);
                                                                LogExceptionDetails("File not extracted: " + destPath + pAttachment, exe);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("File already exists: " + destPath + pAttachment);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (isAlive)
                        db = d.GetNextDatabase();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Extract Graphics Cafe Documents Exception:" + ex);
            }
        }

        private void LogExceptionDetails(string content, Exception exception)
        {
            string exe = Process.GetCurrentProcess().MainModule.FileName;
            string path = Path.GetDirectoryName(exe);
            FileStream fs = new FileStream(path + @"\Logs" + @"\ExceptionLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content + " Message: " + exception?.Message + " Exception: " + exception?.InnerException);
            sw.Flush();
            sw.Close();
        }

        private List<string> ReadSpecificCategories()
        {
            List<string> categories = new List<string>();
            string exe = Process.GetCurrentProcess().MainModule.FileName;
            string path = Path.GetDirectoryName(exe);
            if (File.Exists(path + @"\SpecificCategories.txt"))
            {
                string data = System.IO.File.ReadAllText(path + @"\SpecificCategories.txt");
                categories = data.Split(';').ToList();
                Console.WriteLine("Specific Categories: (" + categories.Count + ") " + data);
            }

            return categories;
        }

        public static void LogDetails(string content)
        {
            string exe = Process.GetCurrentProcess().MainModule.FileName;
            string path = Path.GetDirectoryName(exe);
            FileStream fs = new FileStream(path + @"\Logs" + @"\ServiceLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
    }
}
