class EmployeeICHRAInfo:
    def __init__(self, ssn, dob, home_zip, work_zip, salary, lcsp):
        self.ssn = ssn
        self.dob = dob
        self.home_zip = home_zip
        self.work_zip = work_zip
        self.salary = salary
        self.lcsp = lcsp

    def __str__(self):
        return f"SSN: {self.ssn}, Salary: {self.salary}, DOB: {self.dob}, Work Zip: {self.work_zip}, Home Zip: {self.home_zip}, LCSP: {self.lcsp}"