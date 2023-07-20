using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml.Packaging;
using EmployeeOperations2.Data;
using EmployeeOperations2.Models;
using EmployeeOperations2.Services;
using Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;
using iTextSharp.text;
using PagedList;

namespace EmployeeOperations2.Controllers
{
    public class EmployeesController : Controller
    {
        private EmployeeDBContext db = new EmployeeDBContext();
        private readonly EmailService _emailService = new EmailService();
        List<City> Cities = new List<City>();
        public List<Country> CountriesList; 
        public List<Hobbies> HobbiesOptions; 

        // GET: Employees
        public ActionResult Index(string sortOrder, string CurrentSort, int? page, string searchString)
        {
            int pageSize = 5;
            int pageIndex = 1;
            pageIndex = page.HasValue ? Convert.ToInt32(page) : 1;
            ViewBag.CurrentSort = sortOrder;
            sortOrder = String.IsNullOrEmpty(sortOrder) ? "FirstName" : sortOrder;

            IPagedList<Employee> employees = null;
            switch (sortOrder)
            {
                case "FirstName":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.FirstName).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.FirstName).ToPagedList(pageIndex, pageSize);
                    break;
                case "MiddleName":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.MiddleName).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.MiddleName).ToPagedList(pageIndex, pageSize);
                    break;
                case "LastName":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.LastName).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.LastName).ToPagedList(pageIndex, pageSize);
                    break;
                case "EmailId":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.EmailId).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.EmailId).ToPagedList(pageIndex, pageSize);
                    break;
                case "MobileNumber":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.MobileNumber).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.MobileNumber).ToPagedList(pageIndex, pageSize);
                    break;
                case "Gender":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.Gender).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.Gender).ToPagedList(pageIndex, pageSize);
                    break;
                case "DateOfBirth":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.DateOfBirth).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.DateOfBirth).ToPagedList(pageIndex, pageSize);
                    break;
                case "Country":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.Country).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.Country).ToPagedList(pageIndex, pageSize);
                    break;
                case "City":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.City).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.City).ToPagedList(pageIndex, pageSize);
                    break;
                case "Remarks":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.Remarks).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.Remarks).ToPagedList(pageIndex, pageSize);
                    break;
                case "Hobbies":
                    if (sortOrder.Equals(CurrentSort))
                        employees = db.Employees.OrderByDescending
                                (m => m.Hobbies).ToPagedList(pageIndex, pageSize);
                    else
                        employees = db.Employees.OrderBy
                                (m => m.Hobbies).ToPagedList(pageIndex, pageSize);
                    break;
                case "Default":
                    employees = db.Employees.OrderBy
                        (m => m.FirstName).ToPagedList(pageIndex, pageSize);
                    break;
            }

            if (!string.IsNullOrEmpty(searchString))
            {
                employees = employees.Where(s=> s.FirstName.ToUpper().Contains(searchString.ToUpper())).ToPagedList(pageIndex, pageSize);
            }
                    
            return View(employees);
        }

        private void PopulateCountries()
        {
            CountriesList = new List<Country>
        {
            new Country { Id = 1, Name = "USA" },
            new Country { Id = 2, Name = "India" },
            new Country { Id = 3, Name = "Canada" }
        };
          
        }
        private void PopulateHobbies()
        {
            HobbiesOptions = new List<Hobbies>
            {
                new Hobbies{Id =1, Name="Writing"},
                new Hobbies{Id =2, Name="Playing"},
                new Hobbies{Id =3, Name="Painting"},
                new Hobbies{Id =4, Name="Dancing"},
                new Hobbies{Id =5, Name="Researching"},
                new Hobbies{Id =6, Name="Others"}
            };
        }

        private void PopulateCities()
        {
            //USA
            Cities.Add(new City { Id = 101, Name = "Dallas", countryId = 1 });
            Cities.Add(new City { Id = 102, Name = "NewYork", countryId = 1 });
            Cities.Add(new City { Id = 103, Name = "Seattle", countryId = 1 });

            //India
            Cities.Add(new City { Id = 201, Name = "Hyderabad", countryId = 2 });
            Cities.Add(new City { Id = 202, Name = "NewDelhi", countryId = 2 });
            Cities.Add(new City { Id = 203, Name = "Mumbai", countryId = 2 });

            //Canada
            Cities.Add(new City { Id = 301, Name = "Toronto", countryId = 3 });
            Cities.Add(new City { Id = 302, Name = "Montreal", countryId = 3 });
            Cities.Add(new City { Id = 303, Name = "Vancouver", countryId = 3 });

            //England
            Cities.Add(new City { Id = 401, Name = "London", countryId = 4 });
            Cities.Add(new City { Id = 402, Name = "Manchester", countryId = 4 });
            Cities.Add(new City { Id = 403, Name = "Liverpool", countryId = 4 });

            //Australia
            Cities.Add(new City { Id = 501, Name = "Sydney", countryId = 5 });
            Cities.Add(new City { Id = 502, Name = "Melbourne", countryId = 5 });
            Cities.Add(new City { Id = 503, Name = "Brisbane", countryId = 5 });
        }


        public JsonResult GetCities(int countryId)
        {

            PopulateCities();
            Cities = Cities.Where(c => c.countryId == countryId).ToList();
            return Json(Cities, JsonRequestBehavior.AllowGet);
        }


        // GET: Employees/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // GET: Employees/Create
        public ActionResult Create()
        {
            PopulateCountries();
            PopulateCities();
            PopulateHobbies();

            // Passing data to the view
            ViewBag.Countries = new SelectList(CountriesList, "Id", "Name");
            ViewBag.Cities = new SelectList(Cities, "Id", "Name");
            ViewBag.Hobbies = new SelectList(HobbiesOptions, "Id", "Name"); ;
            return View();
        }



        // POST: Employees/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(Employee employee)
        {
            if (ModelState.IsValid)
            {
                PopulateCountries();
                PopulateCities();
                PopulateHobbies();

                #region Profile Picture
                string photoName = Path.GetFileNameWithoutExtension(employee.PhotoFile.FileName);
                string photoExtension = Path.GetExtension(employee.PhotoFile.FileName);
                photoName = photoName + DateTime.Now.ToString("yyMMdd_mmssfff") + photoExtension;
                employee.PhotoPath = "~/Photos/" + photoName;
                photoName = Path.Combine(Server.MapPath("~/Photos/"), photoName);
                employee.PhotoFile.SaveAs(photoName);
                #endregion

                #region Resume
                string resumeName = Path.GetFileNameWithoutExtension(employee.ResumeFile.FileName);
                string resumeExtension = Path.GetExtension(employee.ResumeFile.FileName);
                resumeName = resumeName + DateTime.Now.ToString("yyMMdd_mmssfff") + resumeExtension;
                employee.ResumePath = "~/Resumes/" + resumeName;
                resumeName = Path.Combine(Server.MapPath("~/Resumes/"), resumeName);
                employee.ResumeFile.SaveAs(resumeName);
                #endregion

                #region Certicicates
                if (employee.CertificateFiles != null && employee.CertificateFiles.Any())
                {
                    foreach (var file in employee.CertificateFiles)
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            string certName = Path.GetFileNameWithoutExtension(file.FileName);
                            string certExtension = Path.GetExtension(file.FileName);
                            certName = certName + DateTime.Now.ToString("yyMMdd_mmssfff") + certExtension;
                            employee.CertificatesPath += "~/Certificates/" + certName+",";
                            certName = Path.Combine(Server.MapPath("~/Certificates/"), certName);
                            file.SaveAs(certName);
                        }
                    }
                }

                employee.CertificatesPath.Remove(employee.CertificatesPath.Length - 1);

                #endregion
                employee.Country = CountriesList.FirstOrDefault(x => x.Id == employee.CountryId).Name;
                employee.City = Cities.FirstOrDefault(x => x.Id == employee.CityId).Name;

                foreach (string hobbieId in employee.HobbiesSelected)
                {
                    employee.Hobbies += HobbiesOptions.FirstOrDefault(h => h.Id == Convert.ToInt32(hobbieId)).Name + ",";
                }
                employee.Hobbies.Remove(employee.Hobbies.Length - 1);

                db.Employees.Add(employee);
                db.SaveChanges();

                SendEmail($"New Entry of {employee.LastName}, {employee.FirstName} is Added.", $"Hi {employee.FirstName}, Welcome to Employee's Operation Corp., Dallas, TX. Thanks.", employee.EmailId, employee.FirstName);
                return RedirectToAction("Index");
            }

            return View(employee);
        }




        [HttpGet]
        public ActionResult DownloadCertificates(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            List<string> certPaths = employee.CertificatesPath.Split(',').ToList();
            int fileNum = 0;
            foreach (string path in certPaths)
            {
                if (path != null && path != string.Empty && path.Length > 3)
                {
                    fileNum++;
                    string filePath = Server.MapPath(path);
                    //byte[] fileBytes;
                    string fileName, contentType, content;
                    //GetDocument(filePath, out fileBytes, out fileName, out contentType);


                    content = System.IO.File.ReadAllText(filePath);

                    string documentsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string fileFullPath = Path.Combine(documentsFolderPath, "Certificate_"+fileNum.ToString()+".pdf");
                    System.IO.File.WriteAllText(fileFullPath, content);

                }
            }
            // Return the file for download
            return View("Details", employee);
        }

        [HttpGet]
        public ActionResult DownloadResume(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            string filePath = Server.MapPath(employee.ResumePath);
            byte[] fileBytes;
            string fileName, contentType;
            GetDocument(filePath, out fileBytes, out fileName, out contentType);
            
            // Return the file for download
            return File(fileBytes, contentType, fileName);
        }

        private static void GetDocument(string filePath, out byte[] fileBytes, out string fileName, out string contentType)
        {
            using (var stream = new MemoryStream())
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open))
                {
                    fileStream.CopyTo(stream);
                }

                fileBytes = stream.ToArray();
            }

            // Set the file name and content type for the download
            string fileExtension = filePath.Substring(filePath.LastIndexOf('.') + 1);
            fileName = string.Empty;
            contentType = string.Empty;
            if (fileExtension == "docx")
            {
                fileName = "Resume.docx";
                contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            }
            else if (fileExtension == "pdf")
            {
                fileName = "Resume.pdf";
                contentType = "application/pdf";
            }
        }

        // GET: Employees/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }

            return View(employee);
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(Employee employee)
        {

            if (ModelState.IsValid)
            {
                #region Profile Picture
                if (employee.PhotoFile != null)
                {
                    string photoName = Path.GetFileNameWithoutExtension(employee.PhotoFile.FileName);
                    string photoExtension = Path.GetExtension(employee.PhotoFile.FileName);
                    photoName = photoName + DateTime.Now.ToString("yymmssfff") + photoExtension;
                    employee.PhotoPath = "~/Photos/" + photoName;
                    photoName = Path.Combine(Server.MapPath("~/Photos/"), photoName);
                    employee.PhotoFile.SaveAs(photoName);
                }
                else
                {
                    Employee original_employee = db.Employees.AsNoTracking().Where(e => e.Id == employee.Id).ToList().FirstOrDefault();
                    employee.PhotoPath = original_employee.PhotoPath;
                    employee.PhotoFile = original_employee.PhotoFile;
                }
                #endregion

                #region Resume
                if (employee.ResumeFile != null)
                {
                    string resumeName = Path.GetFileNameWithoutExtension(employee.ResumeFile.FileName);
                    string resumeExtension = Path.GetExtension(employee.ResumeFile.FileName);
                    resumeName = resumeName + DateTime.Now.ToString("yymmssfff") + resumeExtension;
                    employee.ResumePath = "~/Resumes/" + resumeName;
                    resumeName = Path.Combine(Server.MapPath("~/Resumes/"), resumeName);
                    employee.PhotoFile.SaveAs(resumeName);
                }
                else
                {
                    Employee original_employee = db.Employees.AsNoTracking().Where(e => e.Id == employee.Id).ToList().FirstOrDefault();
                    employee.ResumePath = original_employee.ResumePath;
                    employee.ResumeFile = original_employee.ResumeFile;
                }
                #endregion

                #region Certificates
                if (employee.CertificateFiles != null && employee.CertificateFiles.Any())
                {
                    foreach (var file in employee.CertificateFiles)
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            string certName = Path.GetFileNameWithoutExtension(file.FileName);
                            string certExtension = Path.GetExtension(file.FileName);
                            certName = certName + DateTime.Now.ToString("yyMMdd_mmssfff") + certExtension;
                            employee.CertificatesPath += "~/Certificates/" + certName;
                            certName = Path.Combine(Server.MapPath("~/Certificates/"), certName);
                            file.SaveAs(certName);
                        }
                        else
                        {
                            Employee original_employee = db.Employees.AsNoTracking().Where(e => e.Id == employee.Id).ToList().FirstOrDefault();
                            employee.CertificatesPath = original_employee.CertificatesPath;
                            employee.CertificateFiles = original_employee.CertificateFiles;
                        }
                    }
                }
                Employee original_emp = db.Employees.AsNoTracking().Where(e => e.Id == employee.Id).ToList().FirstOrDefault();
                employee.Hobbies = original_emp.Hobbies;
                #endregion

                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                SendEmail($"{employee.LastName}, {employee.FirstName} has changed.", $"Hi {employee.FirstName}, Your data got changed and updated successfully. Thanks.", employee.EmailId, employee.FirstName);
                return RedirectToAction("Index");
            }
            return View(employee);
        }

        // GET: Employees/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            
            return View(employee);
        }

        // POST: Employees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Employee employee = db.Employees.Find(id);
            db.Employees.Remove(employee);
            db.SaveChanges();
            SendEmail($"{employee.LastName}, {employee.FirstName} has deleted.", $"Hi {employee.FirstName}, Your data got deleted/removed successfully. Thanks.", employee.EmailId, employee.FirstName);
            return RedirectToAction("Index");
        }

        private void SendEmail(string subject, string body, string emailId, string name)
        {
            _emailService.SendEmail(emailId, subject, body, name);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
