using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using ExcelDataReader;
using NickBuhro.Translit;

namespace Add_User
{
    class Program
    {
      static List<ListUsers> listusers = new List<ListUsers>();
    
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                
                Console.ForegroundColor = ConsoleColor.White;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel (*.XLSX)|*.XLSX";
                openFileDialog.ShowDialog();
                string filename = openFileDialog.FileName;
                DataTable dataTable = new DataTable();
                FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);               
                var excelReaderFactory = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                var result =excelReaderFactory.AsDataSet();
                dataTable = result.Tables[0];
                excelReaderFactory.Close();
                foreach (DataRow row in dataTable.Rows)
                {
                    var cells = row.ItemArray;
                    User user = new User();
                    user.FirstName = Convert.ToString(cells[1]);
                    user.LastName = Convert.ToString(cells[0]);
                    user.ThreeName = Convert.ToString(cells[2]);
                    user.Password = GenerationPassword();
                    string Firstname = user.FirstName.Remove(1);
                    string ThreeName = user.ThreeName.Remove(1);
                    string fio = user.LastName + "." + Firstname + ThreeName;
                    string login = Translit(fio);
                    var directoryEntry = new DirectoryEntry("LDAP://Server/CN=Users,DC=NNOV,DC=ru");
                    directoryEntry.Username = "domain\userName";
                    directoryEntry.Password = "Password";
                    Console.WriteLine(String.Format("Добавление пользователя {0} {1} {2} Логин: {3}",user.FirstName,user.LastName,user.ThreeName, login));
                    DirectorySearcher directorySearcher = new DirectorySearcher(directoryEntry);
                    directorySearcher.Filter = ("SAMAccountName=" + login);
                    var resultSearch = directorySearcher.FindOne();
                    if (resultSearch != null)
                    {

                        while (resultSearch != null)
                        {
                            Console.WriteLine(String.Format("Логин {0} уже существует. Введите другой", login));
                            login = Console.ReadLine();
                            directorySearcher.Filter = ("SAMAccountName=" + login);
                            resultSearch = directorySearcher.FindOne();
                        }
                    }
                    DirectoryEntry child = directoryEntry.Children.Add("CN=" + login, "user");                   
                    ///это имя входа
                    child.Properties["SamAccountName"].Value = login;
                    /// имя входа с собакой
                    child.Properties["UserPrincipalName"].Value = login + "@NNOV.ru";
                    ///Полное фио
                    child.Properties["DisplayName"].Value = user.LastName + " " + user.FirstName + " " + user.ThreeName;
                    child.Properties["GivenName"].Value = user.FirstName + " " + user.ThreeName;
                    child.Properties["Sn"].Add(user.ThreeName);
                    child.CommitChanges();
                    DirectoryEntry grp = directoryEntry.Children.Find("CN=DocsVision Users");
                    grp.Invoke("Add", new object[] { "LDAP://Server/CN="+login + ",CN=Users,DC=NNOV,DC=ru" });
                    grp.CommitChanges();
                    child.Invoke("SetPassword", new object[] { user.Password });
                    child.CommitChanges();
                    AccountProperties(login);
                    Logins logins = new Logins()
                    {
                        login = login
                    };
                    AddSpisok(user, logins);
                    
                }
                ListToexcel(listusers);
                Console.WriteLine("Операция выполнена");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
          public static void AccountProperties(string userId)
        {
            try
            {
                PrincipalContext principalContext = new PrincipalContext(ContextType.Machine, "192.168.194.9", "NNOV\\Gusev.la", "Cit934825");
                UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(principalContext, userId);
                userPrincipal.Enabled = true;
                userPrincipal.PasswordNeverExpires = true;
                userPrincipal.UserCannotChangePassword = true;              
                userPrincipal.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
        public static string GenerationPassword ()
        {
            Random random = new Random();
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var chars2 = "abcdefghijklmnopqrstuvwxyz";
            var a = new string(chars.Select(c => chars[random.Next(chars.Length)]).Take(2).ToArray());
            var b = new string(chars.Select(c => chars2[random.Next(chars.Length)]).Take(2).ToArray());
            string chislo =Convert.ToString(random.Next(1000, 9999));
            string result = a + b + chislo;           
                var q = from c in result.ToCharArray()
                        orderby Guid.NewGuid()
                        select c;
                string s = string.Empty;
                foreach (var r in q)
                    s += r;
                return s;
        }
         public static string Translit(string login)
        {
            var latin = Transliteration.CyrillicToLatin(login);
            var a = latin.Replace("`", "");
            return a;
        }

        public static void AddSpisok(User user, Logins login)
        {
            ListUsers listUsers = new ListUsers();
            listUsers.Logins = login;
            listUsers.Users = user;
            listusers.Add(listUsers);
            
           
        }

        public static void ListToexcel(List<ListUsers> lists)
        {
            using (var workbook = new XLWorkbook())
            {
               
                var worksheet = workbook.Worksheets.Add("Пользователи");
                var counter = 1;
                foreach (var item in lists)
                {
                    worksheet.Cell("A" + counter).Value = item.Users.LastName;
                    worksheet.Cell("B" + counter).Value = item.Users.FirstName;
                    worksheet.Cell("C" + counter).Value = item.Users.ThreeName;
                    worksheet.Cell("D" + counter).Value = item.Logins.login;
                    worksheet.Cell("E" + counter).Value = item.Users.Password;
                    ++counter;
                }
                workbook.SaveAs(@"C:\test\users.xlsx");
            }
           

        }


    }
}



    

