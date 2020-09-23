Attribute VB_Name = "modRSUser"
Option Explicit

Public Const sAdministratortitle = "Administrator"
Public Const sEncoderTitle = "Encoder"


Public Const sCanAddUser = "Can Add User"
Public Const sCanEditUser = "Can Edit User"
Public Const sCanDeleteUser = "Can Delete User"
Public Const sCanViewUser = "Can View User"

Public Const sCanAddSchoolYear = "Can Add School Year"
Public Const sCanDeleteSchoolYear = "Can Delete School Year"
Public Const sCanLockUnlockSchoolYear = "Can Lock/Unlock School Year"


Public Const sCanAddDepartment = "Can Add Department"
Public Const sCanEditDepartment = "Can Edit Department"
Public Const sCanDeleteDepartment = "Can Delete Department"

Public Const sCanAddSection = "Can Add Section"
Public Const sCanEditSection = "Can Edit Section"
Public Const sCanDeleteSection = "Can Delete Section"

Public Const sCanAddSectionOffering = "Can Add Section Offering"
Public Const sCanEditSectionOffering = "Can Edit Section Offering"
Public Const sCanDeleteSectionOffering = "Can Delete Section Offering"

Public Const sCanAddTeacher = "Can Add Teacher"
Public Const sCanEditTeacher = "Can Edit Teacher"
Public Const sCanDeleteTeacher = "Can Delete Teacher"

Public Const sCanAddFee = "Can Add Fee"
Public Const sCanEditFee = "Can Edit Fee"
Public Const sCanDeleteFee = "Can Delete Fee"

Public Const sCanAddCashier = "Can Add Cashier"
Public Const sCanEditCashier = "Can Edit Cashier"
Public Const sCanDeleteCashier = "Can Delete Cashier"

Public Const sCanModifyDropped = "Can Add/Remove Dropped Student"

Public Const sCanAddEnrolment = "Can Add Enrolment"
Public Const sCanDeleteEnrolment = "Can Delete Enrolment"
Public Const sCanModifyGraduate = "Can Add/Remove Graduate Student"
Public Const sCanModifyLeaved = "Can Add/Remove Leaving Student"

Public Const sCanAddStudent = "Can Add Student"
Public Const sCanEditStudent = "Can Edit Student"
Public Const sCanDeleteStudent = "Can Delete Student"

Public Const sCanAddCredential = "Can Add Credential"
Public Const sCanEditCredential = "Can Edit Credential"
Public Const sCanDeleteCredential = "Can Delete Credential"

Public Const sCanAddStudentCredential = "Can Add Student Credential"
Public Const sCanDeleteStudentCredential = "Can Delete Student Credential"





Public Const keyUser = "user"


'U S E R
'-----------------------------------------------------
Public Type User
    
    UserName As String
    Password As String
    FullName As String
    UserType As String
    CreationDate As Date
    DateModified As Date
    LastModifiedBy As String
    CreatedBy As String
    
    'misc
    OnLine As Boolean
    
End Type



Public Const CanAddUser = "Can Add User"
Public Const CanEditUser = "Can Edit User"
Public Const CanDeleteUser = "Can Delete User"
Public Const CanClearUserLog = "Can Clear User Log"
    
Public Const CanAddSchoolYear = "Can Add School Year"
Public Const CanEditSchoolYear = "Can Edit School Year"
Public Const CanDeleteSchoolYear = "Can Delete School Year"



