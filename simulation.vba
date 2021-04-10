Function Transactions(r_orders As Range, r_book As Range)
    ReDim orders(r_orders.Rows.Count, r_orders.Columns.Count)
    ReDim book(r_book.Rows.Count, r_book.Columns.Count)
    ReDim trs(r_orders.Rows.Count, 3)
    
    Debug.Print ("----------------------------")
    
    For i = 1 To r_orders.Rows.Count
        For j = 1 To r_orders.Columns.Count
            orders(i - 1, j - 1) = r_orders(i, j)
        Next j
    Next i
    For i = 1 To r_book.Rows.Count
        For j = 1 To r_book.Columns.Count
            book(i - 1, j - 1) = r_book(i, j)
        Next j
    Next i
    
    transaction_count = 0
    
    For i = 1 To r_orders.Rows.Count
        p = orders(i - 1, 0)
        v = orders(i - 1, 1)
        
        best_eladasi = 0
        best_veteli = 0
        
        For k = 1 To r_book.Rows.Count
            If book(k - 1, 2) <> 0 Then
                best_eladasi = book(k - 1, 1)
            End If
        Next k
        
        For k = r_book.Rows.Count To 1 Step -1
            If book(k - 1, 0) <> 0 Then
                best_veteli = book(k - 1, 1)
            End If
        Next k
        
        Debug.Print (transaction_count)
        
        vv = Abs(v)
        
        If v < 0 Then
            ' elad
            If p > best_veteli Then
                ' az legjobb veteli ar kisebb, mint az eladasi ar, nem lesz tranzakcio
                ' bekerul a konyvbe az eladas
                book(20 - p, 2) = book(20 - p, 2) + Abs(v)
            Else
                best_veteli_sor = 20 - best_veteli
                While (best_veteli_sor <= 19 And vv > 0)
                    If book(best_veteli_sor, 0) >= vv Then
                        book(best_veteli_sor, 0) = book(best_veteli_sor, 0) - vv
                        vv = 0
                        trs(transaction_count, 0) = transaction_count + 1
                        trs(transaction_count, 1) = book(best_veteli_sor, 1)
                        trs(transaction_count, 2) = v
                        transaction_count = transaction_count + 1
                    
                    ElseIf book(best_veteli_sor, 0) < vv Then
                        vv = vv - book(best_veteli_sor, 0)
                        
                        If book(best_veteli_sor, 0) > 0 Then
                            trs(transaction_count, 0) = transaction_count + 1
                            trs(transaction_count, 1) = book(best_veteli_sor, 1)
                            trs(transaction_count, 2) = book(best_veteli_sor, 0)
                            
                        End If
                        
                        book(best_veteli_sor, 0) = 0
                        best_veteli_sor = best_veteli_sor + 1
                        
                        
                    End If
                Wend
            End If
        
        ElseIf v > 0 Then
            ' vesz
            If p < best_eladasi Then
                ' az legjobb eladasi ar nagyobb, mint a veteli ar, nem lesz tranzakcio
                ' bekerul a konyvbe a vetel
                book(i - 1, 0) = book(i - 1, 0) + v
            Else
                best_eladasi_sor = 20 - best_eladasi
                While (best_eladasi_sor >= 0 And vv > 0)
                    If book(best_eladasi_sor, 2) >= vv Then
                        book(best_eladasi_sor, 2) = book(best_eladasi_sor, 2) - vv
                        vv = 0
                        trs(transaction_count, 0) = transaction_count + 1
                        trs(transaction_count, 1) = book(best_eladasi_sor, 1)
                        trs(transaction_count, 2) = v
                        transaction_count = transaction_count + 1
                    
                    ElseIf book(best_eladasi_sor, 2) < vv Then
                        vv = vv - book(best_eladasi_sor, 2)
                        
                        If book(best_eladasi_sor, 2) > 0 Then
                            trs(transaction_count, 0) = transaction_count + 1
                            trs(transaction_count, 1) = book(best_eladasi_sor, 1)
                            trs(transaction_count, 2) = book(best_eladasi_sor, 2)
                            
                        End If
                        
                        book(best_eladasi_sor, 2) = 0
                        best_eladasi_sor = best_eladasi_sor - 1
                        
                        
                    End If
                Wend
            End If
        End If
        
    Next i
    
    Transactions = trs
    
End Function



Function CategorizeOrder(r_orders As Range, r_book As Range)
    ReDim orders(r_orders.Rows.Count, r_orders.Columns.Count)
    ReDim book(r_book.Rows.Count, r_book.Columns.Count)
    ReDim res(r_orders.Rows.Count, 1)
    
    For i = 0 To r_orders.Rows.Count - 1
        res(i, 0) = ""
    Next i
    
    Debug.Print ("----------------------------")
    
    For i = 1 To r_orders.Rows.Count
        For j = 1 To r_orders.Columns.Count
            orders(i - 1, j - 1) = r_orders(i, j)
        Next j
    Next i
    For i = 1 To r_book.Rows.Count
        For j = 1 To r_book.Columns.Count
            book(i - 1, j - 1) = r_book(i, j)
        Next j
    Next i
    
    For i = 1 To r_orders.Rows.Count
        p = orders(i - 1, 0)
        v = orders(i - 1, 1)
        
        best_eladasi = 0
        best_veteli = 0
        
        For k = 1 To r_book.Rows.Count
            If book(k - 1, 2) <> 0 Then
                best_eladasi = book(k - 1, 1)
            End If
        Next k
        
        For k = r_book.Rows.Count To 1 Step -1
            If book(k - 1, 0) <> 0 Then
                best_veteli = book(k - 1, 1)
            End If
        Next k
        
        vv = Abs(v)
        
        If p > best_veteli And p < best_eladasi Then
            res(i - 1, 0) = "S"
        End If
        
        If v < 0 Then
            ' elad
            
            If p = best_veteli Then
                res(i - 1, 0) = "P"
            End If
            
            If p > best_veteli Then
                ' az legjobb veteli ar kisebb, mint az eladasi ar, nem lesz tranzakcio
                ' bekerul a konyvbe az eladas
                book(20 - p, 2) = book(20 - p, 2) + Abs(v)
            Else
                best_veteli_sor = 20 - best_veteli
                While (best_veteli_sor <= 19 And vv > 0)
                    If book(best_veteli_sor, 0) >= vv Then
                        book(best_veteli_sor, 0) = book(best_veteli_sor, 0) - vv
                        vv = 0
                    End If
                    If book(best_veteli_sor, 0) < vv Then
                        vv = vv - book(best_veteli_sor, 0)
                        book(best_veteli_sor, 0) = 0
                        best_veteli_sor = best_veteli_sor + 1
                    End If
                Wend
            End If
        End If
        
        If v > 0 Then
            ' vesz
            
            If p = best_eladasi Then
                res(i - 1, 0) = "P"
            End If
            
            If p < best_eladasi Then
                ' az legjobb eladasi ar nagyobb, mint a veteli ar, nem lesz tranzakcio
                ' bekerul a konyvbe a vetel
                book(i - 1, 0) = book(i - 1, 0) + v
            Else
                best_eladasi_sor = 20 - best_eladasi
                While (best_eladasi_sor >= 0 And vv > 0)
                    If book(best_eladasi_sor, 2) >= vv Then
                        book(best_eladasi_sor, 2) = book(best_eladasi_sor, 2) - vv
                        vv = 0
                    End If
                    If book(best_eladasi_sor, 2) < vv Then
                        vv = vv - book(best_eladasi_sor, 2)
                        book(best_eladasi_sor, 2) = 0
                        best_eladasi_sor = best_eladasi_sor - 1
                    End If
                Wend
            End If
        End If
        
    Next i
    
    For i = 0 To r_orders.Rows.Count - 1
        If res(i, 0) = "" Then
            res(i, 0) = "L"
        End If
    Next i
    
    CategorizeOrder = res
    
End Function
