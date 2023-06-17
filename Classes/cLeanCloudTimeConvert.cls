VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLeanCloudTimeConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modTimeZones
' By Chip Pearson, used with permission from www.cpearson.com
' Date: 2-April-2008
' Page Specific URL: www.cpearson.com/Excel/TimeZoneAndDaylightTime.aspx
'
' This module contains functions related to time zones and GMT times.
'   Terms:
'   -------------------------
'   GMT = Greenwich Mean Time. Many applications use the term
'       UTC (Universal Coordinated Time). GMT and UTC are
'       interchangable in meaning,
'   Local Time = The local "wall clock" time of day, that time that
'       you would set a clock to.
'   DST = Daylight Savings Time

'   Functions In This Module:
'   -------------------------
'       ConvertLocalToGMT
'           Converts a local time to GMT. Optionally adjusts for DST.
'       DaylightTime
'           Returns a value indicating (1) DST is in effect, (2) DST is
'           not in effect, or (3) Windows cannot determine whether DST is
'           in effect.
'       GetLocalTimeFromGMT
'           Converts a GMT Time to a Local Time, optionally adjusting for DST.
'       LocalOffsetFromGMT
'           Returns the number of hours/minutes between the local time &GMT,
'           optionally adjusting for DST.
'       SystemTimeToVBTime
'           Converts a SYSTEMTIME structure to a valid VB/VBA date.
'       LocalOffsetFromGMT
'           Returns the number of minutes or hours that are to be added to
'           the local time to get GMT. Optionally adjusts for DST.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Required Types
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum

' Required Windows API Declares
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Sub GetSystemTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME)

Public Function ConvertLocalToGMT(Optional LocalTime As Date, _
    Optional AdjustForDST As Boolean = False) As Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ConvertLocalToGMT
' This converts a local time to GMT. If LocalTime is present, that local
' time is converted to GMT. If LocalTime is omitted, the current time is
' converted from local to GMT. If AdjustForDST is Fasle, no adjustments
' are made to accomodate DST. If AdjustForDST is True, and DST is
' in effect, the time is adjusted for DST by adding
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim T As Date
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    Dim GMT As Date

    If LocalTime <= 0 Then
        T = Now
    Else
        T = LocalTime
    End If
    DST = GetTimeZoneInformation(TZI)
    If AdjustForDST = True Then
        GMT = T + TimeSerial(0, TZI.Bias, 0) + _
                IIf(DST = TIME_ZONE_DAYLIGHT, TimeSerial(0, TZI.DaylightBias, 0), 0)
    Else
        GMT = T + TimeSerial(0, TZI.Bias, 0)
    End If
    ConvertLocalToGMT = GMT
End Function

Public Function GetLocalTimeFromGMT(Optional StartTime As Date) As Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetLocalTimeFromGMT
' This returns the Local Time from a GMT time. If StartDate is present and
' greater than 0, it is assumed to be the GMT from which we will calculate
' Local Time. If StartTime is 0 or omitted, it is assumed to be the GMT
' local time.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim GMT As Date
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    Dim LocalTime As Date

    If StartTime <= 0 Then
        GMT = Now
    Else
        GMT = StartTime
    End If
    DST = GetTimeZoneInformation(TZI)
    LocalTime = GMT - TimeSerial(0, TZI.Bias, 0) + _
            IIf(DST = TIME_ZONE_DAYLIGHT, TimeSerial(1, 0, 0), 0)
    GetLocalTimeFromGMT = LocalTime
End Function

Private Function SystemTimeToVBTime(SysTime As SYSTEMTIME) As Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SystemTimeToVBTime
' This converts a SYSTEMTIME structure to a VB/VBA date value.
' It assumes SysTime is valid -- no error checking is done.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With SysTime
        SystemTimeToVBTime = DateSerial(.wYear, .wMonth, .wDay) + _
                TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function LocalOffsetFromGMT(Optional AsHours As Boolean = False, _
    Optional AdjustForDST As Boolean = False) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LocalOffsetFromGMT
' This returns the amount of time in minutes (if AsHours is omitted or
' false) or hours (if AsHours is True) that should be added to the
' local time to get GMT. If AdjustForDST is missing or false,
' the unmodified difference is returned. (e.g., Kansas City to London
' is 6 hours normally, 5 hours during DST. If AdjustForDST is False,
' the resultif 6 hours. If AdjustForDST is True, the result is 5 hours
' if DST is in effect.)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim TBias As Long
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    DST = GetTimeZoneInformation(TZI)

    If DST = TIME_ZONE_DAYLIGHT Then
        If AdjustForDST = True Then
            TBias = TZI.Bias + TZI.DaylightBias
        Else
            TBias = TZI.Bias
        End If
    Else
        TBias = TZI.Bias
    End If
    If AsHours = True Then
        TBias = TBias / 60
    End If

    LocalOffsetFromGMT = TBias
End Function

Public Function DaylightTime() As TIME_ZONE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DaylightTime
' Returns a value indicating whether the current date is
' in Daylight Time, Standard Time, or that Windows cannot
' deterimine the time status. The result is a member or
' the TIME_ZONE enum.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    DST = GetTimeZoneInformation(TZI)
    DaylightTime = DST
End Function




