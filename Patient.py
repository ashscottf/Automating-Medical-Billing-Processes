class Patient:
    def __init__(self, PatientFirstName, PatientLastName, PatientGender, PatientDOB, PatientPrimaryEmail, PatientPrimaryPhoneNumber, PatientSecondaryPhoneNumber, PatientStreetAddrs, PatientZipCode, PatientFinancalClass, PatientInsId, PatientAccountNumber, PatientCarrier):
        PatientSecondaryPhoneNumbervar = ""
        PatientStreetAddrsVar = ""
        self.PatientFullName = str(PatientLastName).strip() + ", " + str(PatientFirstName).strip()
        self.PatientGender = str(PatientGender).strip().lower()
        self.PatientDOB = str(PatientDOB).strip()
        self.PatientInsID = str(PatientInsId).strip()
        self.PatientPrimaryEmail = str(PatientPrimaryEmail).strip()
        self.PatientPrimaryPhoneNumber = str(PatientPrimaryPhoneNumber).strip()
        self.PatientZipCode = str(PatientZipCode).strip()
        self.PatientAccountNumber = str(PatientAccountNumber).strip()
        self.FinancalClass = str(PatientFinancalClass).strip()
        self.Carrier = str(PatientCarrier).strip()
        
        PatientStreetAddrs = str(PatientStreetAddrs)
        if PatientStreetAddrs[0] == "" or PatientStreetAddrs[0] == " ":
                PatientStreetAddrsVar = PatientStreetAddrs[1:]
        else:
                PatientStreetAddrsVar = str(PatientStreetAddrs)

        if PatientSecondaryPhoneNumber == "" or PatientSecondaryPhoneNumber == " " or PatientSecondaryPhoneNumber == None:
                PatientSecondaryPhoneNumberVar = ""
        else:
                PatientSecondaryPhoneNumbervar = str(PatientSecondaryPhoneNumber).strip()
        self.SecondaryPhoneNumber = PatientSecondaryPhoneNumbervar
        self.PatientStreetAddress = PatientStreetAddrsVar
