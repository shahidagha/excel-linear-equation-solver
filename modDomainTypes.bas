Attribute VB_Name = "modDomainTypes"
Option Explicit

'=============================
' DOMAIN TYPES
'=============================

Public Enum SystemType
    SystemIdentity = 1
    SystemContradictory = 2
    SystemDependent = 3
    SystemIndependent = 4
End Enum

Public Enum EquationStatus
    Normal = 0
    Identity = 1
    Contradiction = 2
End Enum

Public Enum StatusType
    StatusInfo = 0
    StatusSuccess = 1
    StatusError = 2
End Enum

Public Enum FormMode
    ModeIdle = 0
    ModeAdd = 1
    ModeEdit = 2
End Enum


Public Type Surd
    coeff As Long
    radicand As Long
End Type

Public Type FractionSurd
    num As Surd
    den As Surd
End Type

Public Type FractionTerm
    coeff As FractionSurd
    variableID As Integer
End Type

Public Type StandardForm
    aCoeff As FractionSurd
    bCoeff As FractionSurd
    constCoeff As FractionSurd
End Type
