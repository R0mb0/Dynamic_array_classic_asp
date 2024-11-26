<%
class dArray
    'Declaration variables'
    Dim my_array()
    Dim array_dimension
    Dim array_initializated_stauts

    Dim last_index_searched

    ' Initialization and destruction'
	sub class_initialize()
		initialize()
	end sub
	
	sub class_terminate()
		Redim my_array(-1)
        array_dimension = Nothing
        array_initializated_stauts = Nothing
	end sub

    'Function to initialize the array'
    Private Sub initialize()
        array_dimension = 0
        ReDim my_array(array_dimension)
        my_array(array_dimension) = null
        array_initializated_stauts = true
    End Sub

    'Funtion to get array dimension'
    Public Function get_dimension()
        get_dimension = array_dimension
    End Function

    'Funtion to get the entire array'
    Public Function get_array()
        get_array = my_array
    End Function

    'Function to add an element'
    Public Function add_element(element)
        If Not array_initializated_stauts Then 'If the array is not initializated'
            initialize()
            my_array(array_dimension) = element
        ElseIf array_initializated_stauts and array_dimension = 0 and IsNull(my_array(array_dimension)) Then 'If the array has been initializated yet'
            my_array(array_dimension) = element 
        Else'If ther's no special case'
            array_dimension = array_dimension + 1
            Redim Preserve my_array(array_dimension)
            my_array(array_dimension) = element
        End If
    End Function

    'Function to get an element'
    Public Function get_element(idx)
        If idx => 0 and idx <= array_dimension Then
            get_element = my_array(idx)
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - get_element", "Index error")
        End If
    End Function

    'Function to remove last element'
    Public Function remove_last_element()
        array_dimension = array_dimension - 1
        If array_dimension >= 0 Then
            Redim Preserve my_array(array_dimension)
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - remove_last_element", "No element to remove")
        End If
    End Function

    'Funtion to check if an element is in the array'
    Public Function contains(element)
        Dim temp
        Dim temp_index
        temp_index = 0
        For Each temp In my_array
            If temp = element Then
                last_index_searched = temp_index
                contains = true
                Exit Function
            End If
            temp_index = temp_index + 1
        Next
        contains = false
    End Function

    'Function to remove all element occurences from the array'
    Public Function remove_all_occurences(element)
        If contains(element) Then
            Dim temp_array()
            Dim temp_index
            temp_index = 0
            Dim temp
            For Each temp In my_array
                If temp <> element Then
                    ReDim Preserve temp_array(temp_index)
                    temp_array(temp_index) = temp
                    temp_index = temp_index + 1
                End If
            Next
            initialize()
            For Each temp In temp_array
                add_element(temp)
            Next
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - remove_all_occurences", "The element is not present")
        End If
    End Function

    'Function to remove first element occurence from the array'
    Public Function remove_first_occurence(element)
        If contains(element) Then
            Dim temp
            Dim temp_array()
            Dim temp_index
            temp_index = 0
            Dim temp_array_index
            temp_array_index = 0
            For Each temp In my_array
                If temp_index <> last_index_searched Then 
                    Redim Preserve temp_array(temp_array_index)
                    temp_array(temp_array_index) = temp
                    temp_array_index = temp_array_index + 1 
                End If
                temp_index = temp_index + 1
            Next
            initialize()
            For Each temp In temp_array
                add_element(temp)
            Next
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - remove_first_occurence", "The element is not present")
        End If
    End Function

    'Function to remove this element occurence from the array'
    Public Function remove_this_element(idx)
        If idx >= 0 and idx <= array_dimension Then
            my_array(my_index) = Null
            Dim temp
            Dim temp_array()
            Dim temp_index
            temp_index = 0
            For Each temp In my_array
                If Not IsNull(temp) Then
                    ReDim Preserve temp_array(temp_index)
                    temp_array(temp_index) = temp
                    temp_index = temp_index + 1
                End If
            Next
            initialize()
            For Each temp In temp_array
                add_element(temp)
            Next
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - remove_this_element", "Index error")
        End If
    End Function

    'Function to remove these element occurence from the array'
    Public Function remove_these_elements(indices_array)
        Dim my_index
        For Each my_index In indices_array
            If my_index >= 0 and my_index <= array_dimension Then
                my_array(my_index) = Null
            Else
                Call Err.Raise(vbObjectError + 10, "dArray.class - remove_these_elements", "Index error")
            End If
        Next
        Dim temp
        Dim temp_array()
        Dim temp_index
        temp_index = 0
        For Each temp In my_array
            If Not IsNull(temp) Then
                ReDim Preserve temp_array(temp_index)
                temp_array(temp_index) = temp
                temp_index = temp_index + 1
            End If
        Next
        initialize()
        For Each temp In temp_array
            add_element(temp)
        Next
    End Function

    'Function to retrieve the first index of an element in the array'
    Public Function get_first_index_occurence_of(element)
        If contains(element) Then
            get_first_index_occurence_of = last_index_searched
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - get_first_index_occurence_of", "The element is not present")
        End If
    End Function

    'Function to retrieve all indeces of an element in the array'
    Public Function get_all_indeces_occurence_of(element)
        If contains(element) Then
            Dim temp_array()
            Dim temp_array_index
            temp_array_index = 0
            Dim temp_index
            temp_index = 0
            Dim temp
            For Each temp In my_array
                If temp = element Then
                    ReDim Preserve temp_array(temp_array_index)
                    temp_array(temp_array_index) = temp_index
                    temp_array_index = temp_array_index + 1
                End If
                temp_index = temp_index + 1
            Next
            get_all_indeces_occurence_of = temp_array
        Else
            Call Err.Raise(vbObjectError + 10, "dArray.class - get_all_indeces_occurence_of", "The element is not present")
        End If
    End Function

    'Funtion to write the entire array'
    Public Function write_array()
        Dim temp
        Dim temp_index
        temp_index = 0
        For Each temp In my_array
            If temp_index = array_dimension Then
                Response.Write(temp & "<br>")
            Else
                Response.Write(temp & " - ")
            End If
                temp_index = temp_index + 1
        Next
    End Function
end class 
%> 
