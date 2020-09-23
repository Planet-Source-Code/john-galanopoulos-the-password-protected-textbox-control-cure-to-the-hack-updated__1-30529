Attribute VB_Name = "modSubClass"
'* modSubClass.mod - Sub Class Module for SubClass Tracking/Routing
'* ******************************************************
'*
'* Revision:    1.0.0.0
'*
'* ******************************************************
'* Copyright (C) 2001
'* Stephen Kent.
'* ******************************************************
'*
'* Created by:  Stephen Kent
'*
'* Created on:  2001-12-28
'*
'* Project:     SubClass
'*
'* Description: This is the code required to do callback
'*              processing and tracking/routing of messages
'*              to the appropriate subclass object
'*
'* Version control information
'*
'*      Revision:   1.0.0.0
'*      Date:       2001-12-30
'*      Modtime:    09:26 AM
'*      Author:     Stephen Kent
'*
'* ******************************************************
Option Explicit

'Lots of message constants (All I could find that were defined)
Public Const WM_NULL                         As Long = &H0
Public Const WM_GETTEXT                      As Long = &HD
Public Const EM_SETPASSWORDCHAR              As Long = &HCC
'Messages above this number (except WM_APP) are private messages
Public Const WM_USER                         As Long = &H400

Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

'Variable to hold the collection of subclass objects for message routing
Private m_colSubClassObjects As Collection

'This is the call back routine that will replace the window's message routine
Public Function lCallBackProc(ByVal hwnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim oSubClassObject As clsSubClass

    'Check to see if we have our subclass object collection
    If Not (m_colSubClassObjects Is Nothing) Then
        'Loop through all our subclass objects
        For Each oSubClassObject In m_colSubClassObjects
            'Check to see that the object has the window handle and has subclassed it.
            If (oSubClassObject.hwnd = hwnd) And (oSubClassObject.SubClassed) Then
                'Pass the call to the sub class object for processing
                lCallBackProc = oSubClassObject.CallBackProc(hwnd, lMsg, wParam, lParam)
                'We don't want any crashes so only the first object that
                '   subclassed a window will get that message. (Otherwise
                '   it can create an infinite message loop are really screw
                '   things up)
                Exit For
            End If
        Next
    Else
        'No collection so we're stumped - Return False because there's nothing we can do.
        lCallBackProc = False
    End If
End Function

Public Sub AddSubClassObject(oSubClassObject As clsSubClass)
    'Check to see that we have a collection already in existence
    If m_colSubClassObjects Is Nothing Then
        'Nope, Then create one
        Set m_colSubClassObjects = New Collection
    End If
    'Add to the collection
    m_colSubClassObjects.Add oSubClassObject
End Sub

Public Sub RemoveSubClassObject(oSubClassObject As clsSubClass)
    Dim lIndex As Long

    'Check to make sure we have a collection to remove from
    If Not (m_colSubClassObjects Is Nothing) Then
        'loop through all entries until we find a match
        For lIndex = 1 To m_colSubClassObjects.Count
            'Check to see that the hwnd matches and the object is terminating (only terminating when trying to remove itself from the collection)
            If (m_colSubClassObjects(lIndex).hwnd = oSubClassObject.hwnd) And (m_colSubClassObjects(lIndex).Terminating) Then
                'We found a match so remove it and exit the loop
                m_colSubClassObjects.Remove lIndex
                Exit For
            End If
        Next
        'Check to see if there are more objects in collection
        If m_colSubClassObjects.Count = 0 Then
            'No, so destroy the collection
            Set m_colSubClassObjects = Nothing
        End If
    End If
End Sub
