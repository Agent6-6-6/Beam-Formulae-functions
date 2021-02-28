Attribute VB_Name = "Beam_Formulae"
Option Explicit

Private Function getDecPlaces(inputNum As Double) As Long
'returns the number of decimal places within a double

    Dim ndx As Long
    ndx = InStr(1, inputNum, ".")
    If ndx > 0 Then
        getDecPlaces = Len(CStr(inputNum)) - ndx
    End If
End Function

Private Function arr_convert_rng_to_array(arr As Variant)
'function to convert ranges to arrays

    Dim temp As Variant
    'if already an array exit function returning same array
    If IsArray(arr) Then
        arr_convert_rng_to_array = arr
        Exit Function
    End If

    'convert range input into array
    If arr.Columns.count = 1 And arr.Rows.count = 1 Then
        temp = arr.Value2
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = temp
        arr_convert_rng_to_array = arr
    Else
        arr_convert_rng_to_array = arr.Value2
    End If

End Function

Private Function arr_replace_empty(arr As Variant)
'function to replace any empty values in a 2D array with zeros ('0')

    Dim i As Long
    Dim j As Long
    Dim row_num As Long
    Dim col_num As Long

    row_num = UBound(arr, 1)
    col_num = UBound(arr, 2)

    For i = 1 To row_num
        For j = 1 To col_num
            If arr(i, j) = Empty Then arr(i, j) = CDbl(0)
        Next j
    Next i

    'return results
    arr_replace_empty = arr

End Function

Function BEAM_analysis(length_step As Double, L_main As Variant, L_cant As Variant, E_beam As Variant, I_beam As Variant, _
                       Main_UDL As Variant, Main_PT As Variant, Main_PTM As Variant, Cant_UDL As Variant, Cant_PT As Variant, Cant_PTM As Variant, _
                       Optional close_diagram As Boolean = False, Optional moment_shear_only As Boolean = False)
'function to return distance, moment, shear and deflection data for plotting of moment, shear and deflection data

'VARIABLES
'length_step = regularly spaced length along member at which results are reported (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'E_beam = Modulus of elasticity (in MPa)
'I_beam = second moment of area (in mm^4)
'Main_UDL = Cell range or array with UDL load data for the main span
'Main_PT = Cell range or array with point load data for the main span
'Main_PTM = Cell range or array with point moment data for the main span
'Cant_UDL = Cell range or array with UDL load data for the cantilever
'Cant_PT = Cell range or array with point load data for the cantilever
'Cant_PTM = Cell range or array with point moment data for the cantilever
'close_diagram = Optional Boolean value to add extra data points to close off moment and shear diagrams to allow for plotting.
'moment_shear_only = Optional Boolean value to select only returning Moment and Shear

'Beam is analysed as a simply supported span with cantilever to the right hand end
'Loads on the main span are entered relative to the left hand support
'Loads on cantilever are entered relative to the distance from the cantilever end support (right hand end of the the main span)
'NOTE - there are no checks to ensure input locations lie within the defined span lengths.

'Applied load/moment sign convention
'-ve point and UDL loads act downwards
'+ve point moments act clockwise

'initialise variables
    Dim x_distance
    Dim results
    Dim X As Double
    Dim i As Long
    Dim k As Long
    Dim extra_points As Long
    Dim num_columns As Long
    Dim no_cantilever As Boolean
    Dim Main_UDL_loc() As Variant
    Dim Cant_UDL_loc() As Variant
    Dim Main_PT_loc() As Variant
    Dim Cant_PT_loc() As Variant
    Dim Main_PTM_loc() As Variant
    Dim Cant_PTM_loc() As Variant

    If L_cant = 0 Or L_cant = "" Then
        no_cantilever = True
    Else
        no_cantilever = False
    End If

    'convert input to arrays (if cell ranges) & convert empty values to '0' values
    Main_UDL = arr_convert_rng_to_array(Main_UDL)
    Main_UDL = arr_replace_empty(Main_UDL)
    Main_PT = arr_convert_rng_to_array(Main_PT)
    Main_PT = arr_replace_empty(Main_PT)
    Main_PTM = arr_convert_rng_to_array(Main_PTM)
    Main_PTM = arr_replace_empty(Main_PTM)

    If Not no_cantilever Then
        Cant_UDL = arr_convert_rng_to_array(Cant_UDL)
        Cant_UDL = arr_replace_empty(Cant_UDL)
        Cant_PT = arr_convert_rng_to_array(Cant_PT)
        Cant_PT = arr_replace_empty(Cant_PT)
        Cant_PTM = arr_convert_rng_to_array(Cant_PTM)
        Cant_PTM = arr_replace_empty(Cant_PTM)
    End If

    'extract main span start and end points for UDL loading
    ReDim Main_UDL_loc(1 To 2 * UBound(Main_UDL, 1), 1 To 1)
    For i = 1 To UBound(Main_UDL, 1) Step 2
        Main_UDL_loc(i, 1) = Main_UDL(i, 3)    'start location
        Main_UDL_loc(i + 1, 1) = Main_UDL(i, 4)    'end location
    Next i

    'extract main span point load locations
    ReDim Main_PT_loc(1 To UBound(Main_PT, 1), 1 To 1)
    For i = 1 To UBound(Main_PT, 1)
        Main_PT_loc(i, 1) = Main_PT(i, 2)
    Next i

    'extract main span point moment locations
    ReDim Main_PTM_loc(1 To UBound(Main_PTM, 1), 1 To 1)
    For i = 1 To UBound(Main_PTM, 1)
        Main_PTM_loc(i, 1) = Main_PTM(i, 2)
    Next i

    If Not no_cantilever Then
        'extract cantilever span start and end points
        ReDim Cant_UDL_loc(1 To 2 * UBound(Cant_UDL, 1), 1 To 1)
        For i = 1 To UBound(Cant_UDL, 1) Step 2
            Cant_UDL_loc(i, 1) = Cant_UDL(i, 3)    'start location
            Cant_UDL_loc(i + 1, 1) = Cant_UDL(i, 4)    'end location
        Next i

        'extract cantilever span point load locations
        ReDim Cant_PT_loc(1 To UBound(Cant_PT, 1), 1 To 1)
        For i = 1 To UBound(Cant_PT, 1)
            Cant_PT_loc(i, 1) = Cant_PT(i, 2)
        Next i

        'extract cantilever span point moment locations
        ReDim Cant_PTM_loc(1 To UBound(Cant_PTM, 1), 1 To 1)
        For i = 1 To UBound(Cant_PTM, 1)
            Cant_PTM_loc(i, 1) = Cant_PTM(i, 2)
        Next i
    End If

    'generate array of X points
    x_distance = generate_X_array(length_step, L_main, L_cant, _
                                  Main_UDL_loc, Main_PT_loc, Main_PTM_loc, _
                                  Cant_UDL_loc, Cant_PT_loc, Cant_PTM_loc)

    'results array
    '1st column = X DISTANCE
    '2nd column = MOMENT
    '3rd column = SHEAR
    '4th column = DEFLECTION
    If close_diagram Then extra_points = 3
    If moment_shear_only Then
        num_columns = 3
    Else
        num_columns = 4
    End If

    ReDim results(1 To UBound(x_distance) + extra_points, 1 To num_columns)

    For i = 1 To UBound(x_distance)

        X = x_distance(i, 1)
        results(i, 1) = X

        'UDL loads on main span
        For k = 1 To UBound(Main_UDL, 1)
            results(i, 2) = results(i, 2) + BEAM_main_span_UDL_M(Main_UDL(k, 1), Main_UDL(k, 2), X, L_main, Main_UDL(k, 3), Main_UDL(k, 4))
            results(i, 3) = results(i, 3) + BEAM_main_span_UDL_S(Main_UDL(k, 1), Main_UDL(k, 2), X, L_main, Main_UDL(k, 3), Main_UDL(k, 4))
            If Not moment_shear_only Then
                results(i, 4) = results(i, 4) + BEAM_main_span_UDL_D(Main_UDL(k, 1), Main_UDL(k, 2), X, L_main, L_cant, Main_UDL(k, 3), Main_UDL(k, 4), E_beam, I_beam)
            End If
        Next k

        'Point loads on main span
        For k = 1 To UBound(Main_PT)
            results(i, 2) = results(i, 2) + BEAM_main_span_PT_M(Main_PT(k, 1), X, L_main, Main_PT(k, 2))
            results(i, 3) = results(i, 3) + BEAM_main_span_PT_S(Main_PT(k, 1), X, L_main, Main_PT(k, 2))
            If Not moment_shear_only Then
                results(i, 4) = results(i, 4) + BEAM_main_span_PT_D(Main_PT(k, 1), X, L_main, L_cant, Main_PT(k, 2), E_beam, I_beam)
            End If
        Next k

        'Point moments on main span
        For k = 1 To UBound(Main_PTM)
            results(i, 2) = results(i, 2) + BEAM_main_span_PTM_M(Main_PTM(k, 1), X, L_main, Main_PTM(k, 2))
            results(i, 3) = results(i, 3) + BEAM_main_span_PTM_S(Main_PTM(k, 1), X, L_main, Main_PTM(k, 2))
            If Not moment_shear_only Then
                results(i, 4) = results(i, 4) + BEAM_main_span_PTM_D(Main_PTM(k, 1), X, L_main, L_cant, Main_PTM(k, 2), E_beam, I_beam)
            End If
        Next k

        If Not no_cantilever Then
            'UDL loads on cantilever span
            For k = 1 To UBound(Cant_UDL)
                results(i, 2) = results(i, 2) + BEAM_cantilever_UDL_M(Cant_UDL(k, 1), Cant_UDL(k, 2), X, L_main, L_cant, Cant_UDL(k, 3), Cant_UDL(k, 4))
                results(i, 3) = results(i, 3) + BEAM_cantilever_UDL_S(Cant_UDL(k, 1), Cant_UDL(k, 2), X, L_main, L_cant, Cant_UDL(k, 3), Cant_UDL(k, 4))
                If Not moment_shear_only Then
                    results(i, 4) = results(i, 4) + BEAM_cantilever_UDL_D(Cant_UDL(k, 1), Cant_UDL(k, 2), X, L_main, L_cant, Cant_UDL(k, 3), Cant_UDL(k, 4), E_beam, I_beam)
                End If
            Next k

            'Point loads on cantilever span
            For k = 1 To UBound(Cant_PT)
                results(i, 2) = results(i, 2) + BEAM_cantilever_PT_M(Cant_PT(k, 1), X, L_main, Cant_PT(k, 2))
                results(i, 3) = results(i, 3) + BEAM_cantilever_PT_S(Cant_PT(k, 1), X, L_main, Cant_PT(k, 2))
                If Not moment_shear_only Then
                    results(i, 4) = results(i, 4) + BEAM_cantilever_PT_D(Cant_PT(k, 1), X, L_main, L_cant, Cant_PT(k, 2), E_beam, I_beam)
                End If
            Next k

            'Point moments on cantilever span
            For k = 1 To UBound(Cant_PTM)
                results(i, 2) = results(i, 2) + BEAM_cantilever_PTM_M(Cant_PTM(k, 1), X, L_main, Cant_PTM(k, 2))
                results(i, 3) = results(i, 3) + BEAM_cantilever_PTM_S(Cant_PTM(k, 1), X, L_main, Cant_PTM(k, 2))
                If Not moment_shear_only Then
                    results(i, 4) = results(i, 4) + BEAM_cantilever_PTM_D(Cant_PTM(k, 1), X, L_main, L_cant, Cant_PTM(k, 2), E_beam, I_beam)
                End If
            Next k
        End If

    Next i

    If close_diagram = True Then
        'add coordinates to close the geometry/linework of the moment and shear force diagrams for plotting
        results(i, 1) = results(i - 1, 1)
        results(i + 1, 1) = results(1, 1)
        results(i + 2, 1) = results(1, 1)
        results(i, 2) = 0
        results(i + 1, 2) = 0
        results(i + 2, 2) = CVErr(xlErrNA)
        results(i, 3) = 0
        results(i + 1, 3) = 0
        results(i + 2, 3) = results(1, 3)
        If Not moment_shear_only Then
            results(i, 4) = CVErr(xlErrNA)
            results(i + 1, 4) = CVErr(xlErrNA)
            results(i + 2, 4) = CVErr(xlErrNA)
        End If
    End If

    'return results
    BEAM_analysis = results

End Function

Private Function generate_X_array(length_step As Double, L_main As Variant, L_cant As Variant, _
                                  Main_UDL_loc As Variant, Main_PT_loc As Variant, Main_PTM_loc As Variant, _
                                  Cant_UDL_loc As Variant, Cant_PT_loc As Variant, Cant_PTM_loc As Variant)
'function to generate array of sorted and unique X locations, based on start and end points of UDL's and point load locations,
'& beam geometry (supports and start/end of beam), and at a regular spacing defined by 'length_step'

'VARIABLES
'length_step = regularly spaced length along member at which results are reported (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'Main_UDL_loc = array with UDL location load data for the main span
'Main_PT_loc = array with point load location data for the main span
'Main_PTM_loc = array with point moment location data for the main span
'Cant_UDL_loc = array with UDL load location data for the cantilever
'Cant_PT_loc = array with point load location data for the cantilever
'Cant_PTM_loc = array with point moment location data for the cantilever

    Dim num_steps As Long
    Dim num_of_UDL_points As Long
    Dim num_of_point_loads As Long
    Dim no_cantilever As Boolean
    Dim points_of_interest
    Dim i As Long
    Dim k As Long
    Dim tolerance As Double

    'tolerance is the small increment of length to allow for shear to be calculated correctly either side of the locations of applied point loads
    tolerance = 0.0000001

    Dim number_decimals As Long

    number_decimals = getDecPlaces(tolerance) - 1

    If L_cant = 0 Or L_cant = "" Then
        no_cantilever = True
    Else
        no_cantilever = False
    End If

    'number of data steps based on the specified length_step
    num_steps = (L_main + L_cant) / length_step - 1

    '2 times to account for +/- tolerance for working out shears at steps in SFD at point loads
    If Not no_cantilever Then
        num_of_point_loads = 2 * (UBound(Main_PT_loc) + UBound(Cant_PT_loc) + UBound(Main_PTM_loc) + UBound(Cant_PTM_loc))
        num_of_UDL_points = UBound(Main_UDL_loc) + UBound(Cant_UDL_loc)
    Else
        num_of_point_loads = 2 * (UBound(Main_PT_loc) + UBound(Main_PTM_loc))
        num_of_UDL_points = UBound(Main_UDL_loc)
    End If

    'added 3 to put in beginning of main span, end of main span and end of cantilever
    ReDim points_of_interest(1 To num_of_point_loads + num_of_UDL_points + 3 + num_steps, 1 To 1)

    points_of_interest(1, 1) = L_main

    k = 1
    For i = 2 To UBound(Main_PT_loc) * 2 Step 2
        points_of_interest(i, 1) = Main_PT_loc(k, 1) - tolerance
        If points_of_interest(i, 1) < 0 Then points_of_interest(i, 1) = 0
        points_of_interest(i + 1, 1) = Main_PT_loc(k, 1) + tolerance
        If points_of_interest(i + 1, 1) > L_main + L_cant Then points_of_interest(i + 1, 1) = L_main + L_cant
        k = k + 1
    Next i

    k = 1
    For i = i To i + UBound(Main_PTM_loc) * 2 - 1 Step 2
        points_of_interest(i, 1) = Main_PTM_loc(k, 1) - tolerance
        If points_of_interest(i, 1) < 0 Then points_of_interest(i, 1) = 0
        points_of_interest(i + 1, 1) = Main_PTM_loc(k, 1) + tolerance
        If points_of_interest(i + 1, 1) > L_main + L_cant Then points_of_interest(i + 1, 1) = L_main + L_cant
        k = k + 1
    Next i

    If Not no_cantilever Then
        k = 1
        For i = i To i + UBound(Cant_PT_loc) * 2 - 1 Step 2
            points_of_interest(i, 1) = Cant_PT_loc(k, 1) + L_main - tolerance
            If points_of_interest(i, 1) < 0 Then points_of_interest(i, 1) = 0
            points_of_interest(i + 1, 1) = Cant_PT_loc(k, 1) + L_main + tolerance
            If points_of_interest(i + 1, 1) > L_main + L_cant Then points_of_interest(i + 1, 1) = L_main + L_cant
            k = k + 1
        Next i

        k = 1
        For i = i To i + UBound(Cant_PTM_loc) * 2 - 1 Step 2
            points_of_interest(i, 1) = Cant_PTM_loc(k, 1) + L_main - tolerance
            If points_of_interest(i, 1) < 0 Then points_of_interest(i, 1) = 0
            points_of_interest(i + 1, 1) = Cant_PTM_loc(k, 1) + L_main + tolerance
            If points_of_interest(i + 1, 1) > L_main + L_cant Then points_of_interest(i + 1, 1) = L_main + L_cant
            k = k + 1
        Next i
    End If

    k = 1
    For i = i To i + UBound(Main_UDL_loc) - 1
        points_of_interest(i, 1) = Main_UDL_loc(k, 1)
        k = k + 1
    Next i

    If Not no_cantilever Then
        k = 1
        For i = i To i + UBound(Cant_UDL_loc) - 1
            points_of_interest(i, 1) = Cant_UDL_loc(k, 1) + L_main
            k = k + 1
        Next i
    End If

    points_of_interest(i, 1) = 0    'start of beam

    i = i + 1
    For i = i To i + num_steps - 1
        'rounding here deals with numerical precision issues as two returned values are identical but are not treated as such when using WORKSHEET.UNIQUE function
        points_of_interest(i, 1) = Round(points_of_interest(i - 1, 1) + length_step, number_decimals)
    Next i

    points_of_interest(i, 1) = (L_main + L_cant)    'end of beam

    'remove any empty entries (caused by blank entries in input fields)

    For i = 1 To UBound(points_of_interest)
        If points_of_interest(i, 1) = Empty Then points_of_interest(i, 1) = 0
    Next i

    'return results
    With WorksheetFunction
        'sort and return unique values as an array
        generate_X_array = .Sort(.Unique(points_of_interest, False, False), , 1, False)
    End With

End Function

Private Function BEAM_main_span_PT_M(p_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates moment in kNm at point X

'p_1 = Main span point load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point load location from left support (in m)

    Dim moment As Variant

    If X <= L_main Then

        If X <= L_1 Then
            moment = -p_1 * (L_main - L_1) / L_main * (X)
        Else
            moment = -p_1 * L_1 * (L_main - X) / L_main
        End If

    Else

        moment = 0

    End If

    'return results
    BEAM_main_span_PT_M = moment

End Function

Private Function BEAM_main_span_PT_S(p_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates shear in kN at point X

'p_1 = Main span point load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point load location from left support (in m)

    Dim shear As Variant

    If X <= L_main Then

        If X <= L_1 Then
            shear = p_1 * (L_main - L_1) / L_main
        Else
            shear = -p_1 * L_1 / L_main
        End If

    Else

        shear = 0

    End If

    'return results
    BEAM_main_span_PT_S = -shear

End Function

Private Function BEAM_main_span_PT_D(ByVal p_1 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, ByVal L_1 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X

'p_1 = Main span Point Load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point load location from left support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000
'I_beam = second moment of area (in mm^4)

'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000

    'converts kN to N for calculation
    p_1 = p_1 * 1000

    Dim defl As Variant
    Dim pt_defl As Variant
    Dim pt_rotation As Double

    If X <= L_1 Then
        'deflection in main span to the left of the point load
        defl = p_1 * (L_main - L_1) * X * (L_main ^ 2 - (L_main - L_1) ^ 2 - X ^ 2) / (6 * E_beam * I_beam * L_main)

    Else

        If X > L_main Then
            'deflection within cantilever
            defl = -p_1 * L_1 * (L_main - L_1) * (X - L_main) * (L_main + L_1) / (6 * E_beam * I_beam * L_main)
        Else
            'deflection to right of point load location
            defl = p_1 * L_1 * (L_main - X) * (2 * L_main * X - X ^ 2 - L_1 ^ 2) / (6 * E_beam * I_beam * L_main)
        End If

    End If

    'return results
    BEAM_main_span_PT_D = defl

End Function

Private Function BEAM_main_span_PTM_M(m_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates moment in kNm at point X due to point moment

'm_1 = Main span point moment (in kNm)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point moment location from left support (in m)

    Dim moment As Variant

    If X <= L_main Then

        If X <= L_1 Then
            moment = -m_1 * X / L_main
        Else
            moment = m_1 * (L_main - X) / L_main
        End If

    Else

        moment = 0

    End If

    'return results
    BEAM_main_span_PTM_M = moment

End Function

Private Function BEAM_main_span_PTM_S(m_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates shear in kN at point X due to point moment

'm_1 = Main span point moment (in kNm)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point moment location from left support (in m)

    Dim shear As Variant

    If X <= L_main Then

        shear = -m_1 / L_main

    Else

        shear = 0

    End If

    'return results
    BEAM_main_span_PTM_S = shear

End Function

Private Function BEAM_main_span_PTM_D(ByVal m_1 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, ByVal L_1 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X due to point moment

'm_1 = Main span point moment (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point moment location from left support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000
'I_beam = second moment of area (in mm^4)

'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000

    'converts kNm to Nmm for calculation
    m_1 = m_1 * 1000000    'check if required

    Dim defl As Variant
    'Dim pt_defl As Variant
    Dim deflection_from_rotation As Double
    Dim theta_1 As Double
    Dim theta_2 As Double
    Dim R1 As Double

    'support reaction
    R1 = m_1 / L_main

    'support rotation
    theta_1 = m_1 * (3 * (L_main - L_1) ^ 2 - L_main ^ 2) / (6 * E_beam * I_beam * L_main)
    theta_2 = m_1 * (3 * L_1 ^ 2 - L_main ^ 2) / (6 * E_beam * I_beam * L_main)

    'extra cantilever deflection due to support rotation
    deflection_from_rotation = -theta_2 * (X - L_main)

    If X <= L_1 Then
        'deflection in main span to the left of the point moment
        defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam)

    Else

        If X > L_main Then
            'deflection within cantilever
            defl = deflection_from_rotation
        Else
            'deflection to right of point moment location
            defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam) + m_1 * (X - L_1) ^ 2 / (2 * E_beam * I_beam)
        End If

    End If

    'return results
    BEAM_main_span_PTM_D = defl

End Function

Private Function BEAM_cantilever_PT_M(p_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates moment in kNm at point X

'p_1 = Cantilever point load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = Location of point load from right hand support (in m)

    Dim moment As Variant

    If X <= L_main Then
        'X is less than the distance to the right hand support
        moment = p_1 * L_1 * X / L_main
    Else
        If X > L_main + L_1 Then
            'if X is greater than the point at which point load is applied on the beam
            moment = 0
        Else
            'if X is in the cantilever
            moment = p_1 * (L_1 - (X - L_main))
        End If
    End If

    'return results
    BEAM_cantilever_PT_M = moment

End Function

Private Function BEAM_cantilever_PT_S(p_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates shear in kN at point X

'p_1 = Cantilever point load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = Location of point load from right hand support (in m)

    Dim shear As Variant

    If X <= L_main Then
        'shear in main span
        shear = -p_1 * L_1 / L_main
    Else
        If X > L_main + L_1 Then
            shear = 0
        Else
            'shear in cantilever
            shear = p_1
        End If
    End If

    'return results
    BEAM_cantilever_PT_S = -shear

End Function

Private Function BEAM_cantilever_PT_D(ByVal p_1 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, ByVal L_1 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X

'p_1 = Cantilever Point Load (in kN)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = Location of point load from right hand support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000
'I_beam = second moment of area (in mm^4)

'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000

    'converts kN to N for calculation
    p_1 = p_1 * 1000

    Dim defl As Variant
    Dim pt_defl As Variant
    Dim pt_rotation As Double

    If X <= L_main Then
        'deflection in main span
        defl = -p_1 * L_1 * X * (L_main ^ 2 - X ^ 2) / (6 * E_beam * I_beam * L_main)

    Else
        If X > L_main + L_1 Then
            'deflection past point load location

            'variable with deflection at the point load location
            pt_defl = p_1 * L_1 ^ 2 * (L_main + L_1) / (3 * E_beam * I_beam)

            'variable for rotation at point load location (first term is rotation at point load, 2nd term is rotation at support with backspan under moment only
            pt_rotation = -p_1 * (L_1) ^ 2 / (2 * E_beam * I_beam) - p_1 * (L_1) * L_main / (3 * E_beam * I_beam)

            'deflection at location
            defl = pt_defl - pt_rotation * (X - (L_1 + L_main))
        Else
            'deflection between right support and point load
            defl = p_1 * (X - L_main) * (2 * L_1 * L_main + 3 * L_1 * (X - L_main) - (X - L_main) ^ 2) / (6 * E_beam * I_beam)
        End If
    End If

    'return results
    BEAM_cantilever_PT_D = defl

End Function

Private Function BEAM_cantilever_PTM_M(m_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates moment in kNm at point X due to point moment

'm_1 = Cantilever point moment (in kNm)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = Location of point moment from right hand support (in m)

    Dim moment As Variant

    If X <= L_main Then
        'X is less than the distance to the right hand support
        moment = -m_1 * X / L_main
    Else
        If X > L_main + L_1 Then
            'if X is greater than the point at which point moment is applied on the beam
            moment = 0
        Else
            'if X is in the cantilever
            moment = -m_1
        End If
    End If

    'return results
    BEAM_cantilever_PTM_M = moment

End Function

Private Function BEAM_cantilever_PTM_S(m_1 As Variant, X As Variant, L_main As Variant, L_1 As Variant)
'function calculates shear in kN at point X due to point moment

'm_1 = Cantilever point moment (in kNm)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = Location of point moment from right hand support (in m)

    Dim shear As Variant

    If X <= L_main Then
        'shear in main span
        shear = -m_1 / L_main
    Else
        'shear in cantilever
        shear = 0
    End If

    'return results
    BEAM_cantilever_PTM_S = shear

End Function

Private Function BEAM_cantilever_PTM_D(ByVal m_1 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, ByVal L_1 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X

'm_1 = Cantilever Point moment (in kNm)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = Location of point moment from right hand support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000
'I_beam = second moment of area (in mm^4)

'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000

    'converts kNm to Nmm for calculation
    m_1 = m_1 * 1000000

    Dim defl As Variant
    'Dim pt_defl As Variant
    'Dim pt_rotation As Double
    Dim deflection_from_rotation As Double
    Dim theta_1 As Double
    Dim theta_2 As Double
    Dim theta_3 As Double
    Dim R1 As Double

    'support reaction
    R1 = m_1 / L_main

    'support rotation
    theta_1 = m_1 * (-L_main ^ 2) / (6 * E_beam * I_beam * L_main)
    theta_2 = m_1 * (2 * L_main ^ 2) / (6 * E_beam * I_beam * L_main)

    'end of cantilever rotation for fixed case
    theta_3 = m_1 * L_1 / (E_beam * I_beam)

    'extra cantilever deflection due to support rotation
    deflection_from_rotation = -theta_2 * (X - L_main)

    If X <= L_main Then
        'deflection in main span to the left of the point moment
        defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam)

    Else
        If X > L_main + L_1 Then
            'deflection past point moment location
            defl = -theta_3 * ((X - L_main) - L_1 / 2) + deflection_from_rotation
        Else
            'deflection between right support and point load
            defl = -m_1 * (X - L_main) ^ 2 / (2 * E_beam * I_beam) + deflection_from_rotation
        End If
    End If

    'return results
    BEAM_cantilever_PTM_D = defl

End Function

Private Function BEAM_cantilever_UDL_M(w_1 As Variant, w_2 As Variant, X As Variant, L_main As Variant, L_cant As Variant, L_1 As Variant, L_2 As Variant)
'function calculates moment in kNm at point X

'w_1 = Cantilever span partial UDL at start (in kN/m)
'w_2 = Cantilever span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point where partial load starts measured from righthand support (in m)
'L_2 = point where partial load ends measured from righthand support (in m)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end postitions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    Dim moment As Variant
    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim V2 As Double    'shear in cantilever
    Dim R1 As Variant
    Dim R2 As Variant
    Dim M2 As Variant    'moment in cantilever

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_main - L_1)


    'cantilever moment at right hand support
    M2 = -(3 * L_1 + L_w) / 3 * L_w * w_m - L_w ^ 2 * w_2 / 6

    'support reactions
    R1 = M2 / L_main

    'cantilever shear
    V2 = L_w * w_m

    If X <= L_main Then
        'X is less than the distance to the right hand support
        moment = R1 * X
    Else

        If X > L_main + L_cant Then
            'if X is greater than the length of the beam
            moment = 0
        Else
            If X <= L_main + L_1 Then
                'if X is less than the start of the partial UDL
                moment = V2 * (X - L_main) + M2
            Else
                If X >= L_main + L_2 Then
                    'if X is greater than the end of the partial UDL
                    moment = 0
                Else
                    'if X is within the partial UDL
                    moment = V2 * (X - L_main) + M2 - (2 * w_1 + w_x) * (X - L_main - L_1) ^ 2 / 6
                End If
            End If
        End If

    End If

    'return results
    BEAM_cantilever_UDL_M = -moment

End Function

Private Function BEAM_cantilever_UDL_S(w_1 As Variant, w_2 As Variant, X As Variant, L_main As Variant, L_cant As Variant, L_1 As Variant, L_2 As Variant)
'function calculates shear in kN at point X

'w_1 = Cantilever span partial UDL at start (in kN/m)
'w_2 = Cantilever span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point where partial load starts measured from righthand support (in m)
'L_2 = point where partial load ends measured from righthand support (in m)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end postitions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    Dim shear As Variant
    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim V2 As Double    'shear in cantilever
    Dim R1 As Variant
    Dim R2 As Variant
    Dim M2 As Variant    'moment in cantilever

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_main - L_1)

    'cantilever moment at right hand support
    M2 = -(3 * L_1 + L_w) / 3 * L_w * w_m - L_w ^ 2 * w_2 / 6

    'support reactions
    R1 = M2 / L_main

    'cantilever shear
    V2 = L_w * w_m

    If X <= L_main Then
        'shear in main span
        shear = R1

    Else
        If X > L_main + L_cant Then
            shear = 0
        Else
            If X <= L_main + L_1 Then
                'if X is less than the start of the partial UDL
                shear = V2
            Else
                If X >= L_main + L_2 Then
                    'if X is greater than the end of the partial UDL
                    shear = 0
                Else
                    'if X is within the partial UDL
                    shear = V2 - (w_1 + w_x) * (X - L_main - L_1) / 2
                End If
            End If
        End If

    End If

    'return results
    BEAM_cantilever_UDL_S = -shear

End Function

Private Function BEAM_cantilever_UDL_D(ByVal w_1 As Variant, ByVal w_2 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, _
                                       ByVal L_1 As Variant, ByVal L_2 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X

'w_1 = Cantilever span partial UDL at start (in kN/m)
'w_2 = Cantilever span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point where partial load starts measured from righthand support (in m)
'L_2 = point where partial load ends measured from righthand support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000 or 205000
'I = second moment of area (in mm^4)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end postitions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000
    L_2 = L_2 * 1000

    Dim defl As Variant
    Dim theta_1 As Double
    Dim theta_2 As Double
    Dim theta_3 As Double
    Dim extra_deflection_from_rotation
    Dim x_dash
    Dim a
    Dim b
    Dim c
    Dim f
    Dim end_defl

    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim V2 As Double    'shear in cantilever
    Dim R1 As Variant
    Dim R2 As Variant
    Dim M2 As Variant    'moment in cantilever

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_main - L_1)

    'cantilever moment at right hand support
    M2 = -(3 * L_1 + L_w) / 3 * L_w * w_m - L_w ^ 2 * w_2 / 6

    'support reactions
    R1 = M2 / L_main

    'cantilever shear
    V2 = L_w * w_m

    'support rotations
    theta_1 = -M2 * L_main ^ 2 / (6 * E_beam * I_beam * L_main)
    theta_2 = M2 * (2 * L_main ^ 2) / (6 * E_beam * I_beam * L_main)
    theta_3 = -L_w * ((2 * L_1 ^ 2 + (L_1 + L_2) ^ 2) * w_m + (L_1 + L_2) * L_w * w_2) / (12 * E_beam * I_beam)

    'extra cantilever deflection due to support rotation
    extra_deflection_from_rotation = -theta_2 * (X - L_main)

    If X <= L_main Then
        'deflection in main span
        defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam)
    Else
        If X <= L_main + L_1 Then
            'if X is less than the start of the partial UDL
            defl = -V2 * (X - L_main) ^ 3 / (6 * E_beam * I_beam) - M2 * (X - L_main) ^ 2 / (2 * E_beam * I_beam) + extra_deflection_from_rotation
        Else
            If X >= L_main + L_2 Then
                'if X is greater than the end of the partial UDL
                defl = -V2 * (L_2) ^ 3 / (6 * E_beam * I_beam) - M2 * (L_2) ^ 2 / (2 * E_beam * I_beam) + (4 * w_1 + w_2) * (L_w) ^ 4 / (120 * E_beam * I_beam) - _
                       theta_3 * (X - L_main - L_2) + extra_deflection_from_rotation
            Else
                'if X is within the partial UDL
                defl = -V2 * (X - L_main) ^ 3 / (6 * E_beam * I_beam) - M2 * (X - L_main) ^ 2 / (2 * E_beam * I_beam) + _
                       (4 * w_1 + w_x) * (X - L_main - L_1) ^ 4 / (120 * E_beam * I_beam) + extra_deflection_from_rotation
            End If
        End If
    End If

    'return results
    BEAM_cantilever_UDL_D = defl

End Function

Private Function BEAM_main_span_UDL_M(w_1 As Variant, w_2 As Variant, X As Variant, L_main As Variant, L_1 As Variant, L_2 As Variant)
'function calculates moment in kNm at point X

'w_1 = Main span partial UDL at start (in kN/m)
'w_2 = Main span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point where partial load starts measured from lefthand support (in m)
'L_2 = point where partial load ends measured from lefthand support (in m)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end positions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    Dim moment As Variant
    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim R1 As Variant
    Dim R2 As Variant

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_1)

    'support reactions
    R1 = L_w * ((6 * w_m * (L_main - L_2) + (2 * w_1 + w_2) * L_w) / (6 * L_main))
    R2 = L_w * ((6 * w_m * (L_2) - (2 * w_1 + w_2) * L_w) / (6 * L_main))

    If X <= L_main Then
        If X < L_1 Then
            ' case where X lies before the partial load position
            moment = R1 * X
        Else
            If X > L_2 Then
                ' case where X lies after the partial load position
                moment = R2 * (L_main - X)
            Else
                'case where X lies within partial UDL
                moment = R1 * X - ((2 * w_1 + w_x) * (X - L_1) ^ 2) / 6
            End If
        End If
    Else
        moment = 0
    End If

    'return results
    BEAM_main_span_UDL_M = -moment

End Function

Private Function BEAM_main_span_UDL_S(w_1 As Variant, w_2 As Variant, X As Variant, L_main As Variant, L_1 As Variant, L_2 As Variant)
'function calculates shear in kN at point X

'w_1 = Main span partial UDL at start (in kN/m)
'w_2 = Main span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_1 = point where partial load starts measured from lefthand support (in m)
'L_2 = point where partial load ends measured from lefthand support (in m)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end postitions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    Dim shear As Variant
    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim R1 As Variant
    Dim R2 As Variant

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_1)

    'support reactions
    R1 = L_w * ((6 * w_m * (L_main - L_2) + (2 * w_1 + w_2) * L_w) / (6 * L_main))
    R2 = L_w * ((6 * w_m * (L_2) - (2 * w_1 + w_2) * L_w) / (6 * L_main))

    If X <= L_main Then
        If X < L_1 Then
            ' case where X lies before the partial load position
            shear = R1
        Else
            If X > L_2 Then
                ' case where X lies after the partial load position
                shear = -R2
            Else
                'case where X lies within partial UDL
                shear = R1 - (w_1 + w_x) * (X - L_1) / 2
            End If
        End If
    Else
        shear = 0
    End If

    'return results
    BEAM_main_span_UDL_S = -shear

End Function

Private Function BEAM_main_span_UDL_D(ByVal w_1 As Variant, ByVal w_2 As Variant, ByVal X As Variant, ByVal L_main As Variant, ByVal L_cant As Variant, _
                                      ByVal L_1 As Variant, ByVal L_2 As Variant, E_beam As Variant, I_beam As Variant)
'function calculates deflection in mm at point X

'w_1 = Main span partial UDL at start (in kN/m)
'w_2 = Main span partial UDL at end (in kN/m)
'X = point being measured from lefthand support (in m)
'L_main = Main span (backspan) length (in m)
'L_cant = Cantilever span length (in m)
'L_1 = point where partial load starts measured from lefthand support (in m)
'L_2 = point where partial load ends measured from lefthand support (in m)
'E_beam = Modulus of elasticity (in MPa) i.e. steel = 200000 or 205000
'I_beam = second moment of area (in mm^4)

'check for zero length loads
    If L_1 = L_2 Then Exit Function

    'check start & end postitions and swap if required
    If L_1 > L_2 Then
        Dim temp
        temp = L_1
        L_1 = L_2
        L_2 = temp
        temp = w_1
        w_1 = w_2
        w_2 = temp
    End If

    'convert lengths to mm for calculation
    X = X * 1000
    L_main = L_main * 1000
    L_cant = L_cant * 1000
    L_1 = L_1 * 1000
    L_2 = L_2 * 1000

    Dim defl As Variant
    Dim L_w As Double
    Dim w_m As Double
    Dim w_x As Double
    Dim R1 As Variant
    Dim R2 As Variant
    Dim theta_1 As Double
    Dim theta_2 As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double
    Dim s4 As Double

    L_w = L_2 - L_1
    w_m = (w_1 + w_2) / 2
    w_x = w_1 + ((w_2 - w_1) / L_w) * (X - L_1)

    'support reactions
    R1 = L_w * ((6 * w_m * (L_main - L_2) + (2 * w_1 + w_2) * L_w) / (6 * L_main))
    R2 = L_w * ((6 * w_m * (L_2) - (2 * w_1 + w_2) * L_w) / (6 * L_main))

    'intermediate parameters
    s1 = 20 * L_1 ^ 2 * (L_1 - 3 * L_main) + 20 * L_w * L_1 * (L_1 - 2 * L_main) + 10 * L_w ^ 2 * (L_1 - L_main) + 2 * L_w ^ 3
    s2 = 10 * L_w * L_1 * (L_1 - 2 * L_main) + 10 * L_w ^ 2 * (L_1 - L_main) + 3 * L_w ^ 3
    s3 = 20 * L_1 ^ 3 + 20 * L_w * L_1 ^ 2 + 10 * L_w ^ 2 * L_1 + 2 * L_w ^ 3
    s4 = 10 * L_w * L_1 ^ 2 + 10 * L_w ^ 2 * L_1 + 3 * L_w ^ 3

    'support rotations
    theta_1 = -R2 * L_main ^ 2 / (3 * E_beam * I_beam) - L_w * (s1 * w_m + s2 * w_2) / (120 * E_beam * I_beam * L_main)
    theta_2 = R2 * L_main ^ 2 / (6 * E_beam * I_beam) - L_w * (s3 * w_m + s4 * w_2) / (120 * E_beam * I_beam * L_main)

    If X <= L_main Then
        If X <= L_1 Then
            ' case where X lies before the partial load position
            defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam)
        Else
            If X > L_2 Then
                ' case where X lies after the partial load position
                defl = theta_2 * (L_main - X) - R2 * (L_main - X) ^ 3 / (6 * E_beam * I_beam)
            Else
                'case where X lies within partial UDL
                defl = -theta_1 * X - R1 * X ^ 3 / (6 * E_beam * I_beam) + ((4 * w_1 + w_x) * (X - L_1) ^ 4) / (120 * E_beam * I_beam)
            End If
        End If
    Else
        'deflection in the cantilever
        defl = -theta_2 * (X - L_main)
    End If

    'return results
    BEAM_main_span_UDL_D = defl

End Function
