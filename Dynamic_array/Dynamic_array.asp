<%
'------- Variable part--------'

'Declaration variables'
Dim my_array()
Dim array_dimension
Dim array_initializated_stauts

Dim last_index_searched

'Initialize variables'
array_initializated_stauts = false
array_dimension = null

'-------- Functions part--------'

'Funtion to get array status'
Function is_array_initializated()
  is_array_initializated = array_initializated_stauts
End Function

'Function to initialize the array'
Function initialize_array()
  array_dimension = 0
  ReDim my_array(array_dimension)
  my_array(array_dimension) = null
  array_initializated_stauts = true
End Function

'Funtion to get array dimension'
Function get_array_dimension()
  get_array_dimension = array_dimension
End Function

'Funtion to get the entire array'
Function get_array()
  get_array = my_array
End Function

'Function to add an element'
Function add_element_to_array(ByVal element)
  If Not array_initializated_stauts Then 'If the array is not initializated'
    initialize_array()
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
Function get_element_from_array(ByVal idx)
  If idx => 0 and idx <= array_dimension Then
    get_element_from_array = my_array(idx)
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - get_element_from_array", "Index error: "&idx&"")
  End If
End Function

'Function to remove last element'
Function remove_last_element_from_array()
  array_dimension = array_dimension - 1
  If array_dimension >= 0 Then
    Redim Preserve my_array(array_dimension)
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - remove_last_element_from_array", "No element to remove")
  End If
End Function

'Funtion to check if an element is in the array'
Function array_contains(ByVal element)
  Dim temp
  Dim temp_index
  temp_index = 0
  For Each temp In my_array
    If temp = element Then
      last_index_searched = temp_index
      array_contains = true
      Exit Function
    End If
    temp_index = temp_index + 1
  Next
  array_contains = false
End Function

'Function to remove all element occurences from the array'
Function remove_all_occurences_from_array(ByVal element)
  If array_contains(element) Then
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
    initialize_array()
    For Each temp In temp_array
      add_element_to_array(temp)
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - remove_all_occurences_from_array", "The element "&element&" is not present")
  End If
End Function

'Function to remove first element occurence from the array'
Function remove_first_occurence_from_array(ByVal element)
  If array_contains(element) Then
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
    initialize_array()
    For Each temp In temp_array
      add_element_to_array(temp)
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - remove_first_occurence_from_array", "The element "&element&" is not present")
  End If
End Function

'Function to remove this element occurence from the array'
Function remove_this_element_from_array(ByVal idx)
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
    initialize_array()
    For Each temp In temp_array
      add_element_to_array(temp)
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - remove_this_element_from_array", "Index error: "&idx&"")
  End If
End Function

'Function to remove these element occurence from the array'
Function remove_these_elements_from_array(ByVal indices_array)
  Dim my_index
  For Each my_index In indices_array
    If my_index >= 0 and my_index <= array_dimension Then
      my_array(my_index) = Null
    Else
      Call Err.Raise(vbObjectError + 10, "Dynamic_array - remove_these_elements_from_array", "Index error: "&my_index&"")
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
  initialize_array()
  For Each temp In temp_array
    add_element_to_array(temp)
  Next
End Function

'Function to retrieve the first index of an element in the array'
Function from_array_get_first_index_occurence_of(ByVal element)
  If array_contains(element) Then
    from_array_get_first_index_occurence_of = last_index_searched
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - from_array_get_first_index_occurence_of", "The element "&element&" is not present")
  End If
End Function

'Function to retrieve all indeces of an element in the array'
Function from_array_get_all_indeces_occurence_of(ByVal element)
  If array_contains(element) Then
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
    from_array_get_all_indeces_occurence_of = temp_array
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_array - from_array_get_all_indeces_occurence_of", "The element "&element&" is not present")
  End If
End Function

'Funtion to write the entire array'
Function write_array()
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
%>
