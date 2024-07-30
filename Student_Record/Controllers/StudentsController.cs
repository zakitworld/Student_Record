using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Student_Record.Data;
using Student_Record.Models;
using X.PagedList;

namespace Student_Record.Controllers
{
    public class StudentsController : Controller
    {
        private readonly Student_RecordDbContext _context;

        public StudentsController(Student_RecordDbContext context)
        {
            _context = context;
        }

        public IActionResult ImportExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["ErrorMessage"] = "Please upload a valid Excel file.";
                return RedirectToAction(nameof(Index));
            }

            if (!file.FileName.EndsWith(".xlsx"))
            {
                TempData["ErrorMessage"] = "The file format is not supported. Please upload an Excel file with .xlsx extension.";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        if (worksheet == null)
                        {
                            TempData["ErrorMessage"] = "The Excel file is empty or the first worksheet is missing.";
                            return RedirectToAction(nameof(Index));
                        }

                        var rowCount = worksheet.Dimension.Rows;
                        var colCount = worksheet.Dimension.Columns;

                        if (colCount < 8)
                        {
                            TempData["ErrorMessage"] = "The Excel file is missing some required columns.";
                            return RedirectToAction(nameof(Index));
                        }

                        var students = new List<Students>();

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var studentName = worksheet.Cells[row, 1].Text.Trim();
                            var fatherName = worksheet.Cells[row, 2].Text.Trim();
                            var fatherContact = worksheet.Cells[row, 3].Text.Trim();
                            var motherName = worksheet.Cells[row, 4].Text.Trim();
                            var motherContact = worksheet.Cells[row, 5].Text.Trim();
                            var dateOfBirthText = worksheet.Cells[row, 6].Text.Trim();
                            var gender = worksheet.Cells[row, 7].Text.Trim();
                            var programmeEnrolled = worksheet.Cells[row, 8].Text.Trim();

                            if (string.IsNullOrEmpty(studentName) || string.IsNullOrEmpty(fatherName) || string.IsNullOrEmpty(fatherContact) ||
                                string.IsNullOrEmpty(motherName) || string.IsNullOrEmpty(motherContact) || string.IsNullOrEmpty(dateOfBirthText) ||
                                string.IsNullOrEmpty(gender) || string.IsNullOrEmpty(programmeEnrolled))
                            {
                                TempData["ErrorMessage"] = $"Row {row} has missing values. Please ensure all required fields are filled.";
                                return RedirectToAction(nameof(Index));
                            }

                            if (!DateTime.TryParse(dateOfBirthText, out var dateOfBirth))
                            {
                                TempData["ErrorMessage"] = $"Row {row} has an invalid date of birth. Please ensure the date is in the correct format.";
                                return RedirectToAction(nameof(Index));
                            }

                            students.Add(new Students
                            {
                                StudentName = studentName,
                                FatherName = fatherName,
                                FatherContact = fatherContact,
                                MotherName = motherName,
                                MotherContact = motherContact,
                                DateOfBirth = dateOfBirth,
                                Gender = gender,
                                ProgrammeEnrolled = programmeEnrolled
                            });
                        }

                        _context.Students.AddRange(students);
                        await _context.SaveChangesAsync();
                        TempData["SuccessMessage"] = "Excel file imported successfully!";
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"An error occurred while processing the file: {ex.Message}";
            }

            return RedirectToAction(nameof(Index));
        }

        public IActionResult DownloadExcel()
        {
            var students = _context.Students.ToList(); // Retrieve all students from database

            // Create Excel package using EPPlus
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Students");

                // Header row
                worksheet.Cells["A1"].Value = "Student Name";
                worksheet.Cells["B1"].Value = "Father's Name";
                worksheet.Cells["C1"].Value = "Father's Contact";
                worksheet.Cells["D1"].Value = "Mother's Name";
                worksheet.Cells["E1"].Value = "Mother's Contact";
                worksheet.Cells["F1"].Value = "Date of Birth";
                worksheet.Cells["G1"].Value = "Gender";
                worksheet.Cells["H1"].Value = "Programme Enrolled";

                // Data rows
                int row = 2;
                foreach (var student in students)
                {
                    worksheet.Cells[$"A{row}"].Value = student.StudentName;
                    worksheet.Cells[$"B{row}"].Value = student.FatherName;
                    worksheet.Cells[$"C{row}"].Value = student.FatherContact;
                    worksheet.Cells[$"D{row}"].Value = student.MotherName;
                    worksheet.Cells[$"E{row}"].Value = student.MotherContact;
                    worksheet.Cells[$"F{row}"].Value = student.DateOfBirth.ToString("yyyy-MM-dd");
                    worksheet.Cells[$"G{row}"].Value = student.Gender;
                    worksheet.Cells[$"H{row}"].Value = student.ProgrammeEnrolled;
                    row++;
                }

                // Prepare response stream
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                // Set content type and file name
                string excelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string excelFileName = $"StudentRecords_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                return File(stream, excelContentType, excelFileName);
            }
        }

        // GET: Students
        public async Task<IActionResult> Index(string searchString, int? page)
        {
            if (_context.Students == null)
            {
                return Problem("Entity set 'Student_RecordDbContext.Students' is null.");
            }

            var students = from s in _context.Students select s;

            if (!String.IsNullOrEmpty(searchString))
            {
                students = students.Where(s => s.StudentName!.Contains(searchString));
            }

            int pageSize = 10;
            int pageNumber = page ?? 1;
            var pagedStudents = await students.ToPagedListAsync(pageNumber, pageSize);

            return View(pagedStudents);
        }

        // GET: Students/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var students = await _context.Students
                .FirstOrDefaultAsync(m => m.Id == id);
            if (students == null)
            {
                return NotFound();
            }

            return View(students);
        }

        // GET: Students/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Students/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,StudentName,FatherName,FatherContact,MotherName,MotherContact,DateOfBirth,Gender,ProgrammeEnrolled")] Students students)
        {
            if (ModelState.IsValid)
            {
                _context.Add(students);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Student created successfully!";
                return RedirectToAction(nameof(Index));
            }
            return View(students);
        }

        // GET: Students/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var students = await _context.Students.FindAsync(id);
            if (students == null)
            {
                return NotFound();
            }
            return View(students);
        }

        // POST: Students/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,StudentName,FatherName,FatherContact,MotherName,MotherContact,DateOfBirth,Gender,ProgrammeEnrolled")] Students students)
        {
            if (id != students.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(students);
                    await _context.SaveChangesAsync();
                    TempData["SuccessMessage"] = "Student updated successfully!";
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!StudentsExists(students.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(students);
        }

        // GET: Students/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var students = await _context.Students
                .FirstOrDefaultAsync(m => m.Id == id);
            if (students == null)
            {
                return NotFound();
            }

            return View(students);
        }

        // POST: Students/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var students = await _context.Students.FindAsync(id);
            if (students != null)
            {
                _context.Students.Remove(students);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Student deleted successfully!";
            }

            return RedirectToAction(nameof(Index));
        }

        private bool StudentsExists(int id)
        {
            return _context.Students.Any(e => e.Id == id);
        }
    }
}
