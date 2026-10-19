Attribute VB_Name = "modProducts"
Option Explicit

Public Type Product
    Name As String * 12
    Price As Currency
    Active As Boolean
End Type

Public Products() As Product

' Reads the price list (a Long count, then the records).
Public Sub LoadProducts(ByVal path As String)
    Dim f As Integer, n As Long, i As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    Get #f, , n
    ReDim Products(1 To n)
    For i = 1 To n
        Get #f, , Products(i)
    Next
    Close #f
End Sub

' The price of a product by name.
Public Function FindPrice(ByVal productName As String) As Currency
    Dim i As Long
    For i = 1 To UBound(Products)
        If RTrim$(Products(i).Name) = productName Then
            FindPrice = Products(i).Price
            Exit Function
        End If
    Next
End Function
