<%

'------- Variable part--------'

'Declaration variables'
Dim last_index_searched

'---- Functions ----'

'Function to initialize the array'
Function get_initializated_dynamic_array()
  Dim temp
  temp = Array()
  Redim temp(0)
  temp(0) = NULL
  get_initializated_dynamic_array = temp
End Function

'Function to add an element'
Function add_element_to_dynamic_array(ByRef my_array,ByVal element)
  Dim temp
  temp = UBound(my_array)
  If temp = 0 and IsNull(my_array(0)) Then 'If the array has been just initializated'
    my_array(temp)= element 
  Else'If ther's no special case'
    temp = temp + 1
  Redim Preserve my_array(temp)
  my_array(temp) = element
End If
End Function

'Function to get an element'
Function get_element_from_dynamic_array(ByRef my_array,ByVal idx)
  Dim temp
  temp = UBound(my_array)
  If idx >= 0 and idx <= temp Then
    get_element_from_dynamic_array = my_array(idx)
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - get_element_from_dynamic_array", "Index error: "&idx&"")
  End If
End Function

'Function to remove last element'
Function remove_last_element_from_dynamic_array(ByRef my_array)
  Dim temp
  temp = UBound(my_array)
  temp = temp - 1
  If temp >= 0 Then
    Redim Preserve my_array(temp)
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - remove_last_element_from_dynamic_array", "No element to remove")
  End If
End Function

'Funtion to check if an element is in the dynamic array'
Function dynamic_array_contains(ByRef my_array,ByVal element)
  Dim temp
  Dim temp_index
  temp_index = 0
  For Each temp In my_array
    If temp = element Then
      last_index_searched = temp_index
      dynamic_array_contains = true
      Exit Function
    End If
    temp_index = temp_index + 1
  Next
  dynamic_array_contains = false
End Function

'Function to remove all element occurences from the dynamic array'
Function remove_all_occurences_from_dynamic_array(ByRef my_array,ByVal element)
  If dynamic_array_contains(my_array,element) Then
    Dim temp_array
    temp_array = Array()
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
  my_array = get_initializated_dynamic_array()
  For Each temp In temp_array
    add_element_to_dynamic_array my_array,temp
  Next
  Else
  Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - remove_all_occurences_from_dynamic_array", "No element "&element&" to remove")
  End If
End Function

'Function to remove first element occurence from the array'
Function remove_first_occurence_from_dynamic_array(ByRef my_array,ByVal element)
  If dynamic_array_contains(my_array,element) Then
    Dim temp_array()
    Dim temp_index
    temp_index = 0
    Dim temp_array_index
    temp_array_index = 0
    Dim temp
    For Each temp In my_array
      If temp_index <> last_index_searched Then
          Redim Preserve temp_array(temp_array_index)
          temp_array(temp_array_index) = temp
          temp_array_index = temp_array_index + 1
      End If
      temp_index = temp_index + 1
    Next
    my_array = get_initializated_dynamic_array()
    For Each temp In temp_array
      add_element_to_dynamic_array my_array,temp
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - remove_first_occurence_from_dynamic_array", "No element "&element&" to remove")
  End If
End Function

'Function to remove these element occurence from the dynamic array'
Function remove_these_elements_from_dynamic_array(ByRef my_array,ByVal indices_array)
  Dim my_index
  Dim array_dimension
  array_dimension = UBound(my_array)
  For Each my_index In indices_array
    If my_index >= 0 and my_index <= array_dimension Then
      my_array(my_index) = Null
    Else
      Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - remove_these_elements_from_dynamic_array", "Index error: "&my_index&"")
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
  my_array = get_initializated_dynamic_array()
  For Each temp In temp_array
    add_element_to_dynamic_array my_array,temp
  Next
End Function

'Function to remove this element occurence from the dynamic array'
Function remove_this_element_from_dynamic_array(ByRef my_array,ByVal idx)
  Dim array_dimension
  array_dimension = UBound(my_array)
  If idx >= 0 and idx <= array_dimension Then
    my_array(idx) = Null
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - remove_this_element_from_dynamic_array", "Index error: "&idx&"")
  End If
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
  my_array = get_initializated_dynamic_array()
  For Each temp In temp_array
    add_element_to_dynamic_array my_array,temp
  Next
End Function

'Function to retrieve the first index of an element in the dynamic array'
Function from_dynamic_array_get_first_index_occurence_of(ByRef my_array,ByVal element)
  If dynamic_array_contains(my_array,element) Then
    from_dynamic_array_get_first_index_occurence_of = last_index_searched
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - from_dynamic_array_get_first_index_occurence_of", "The element "&element&" is not present")
  End If
End Function

'Function to retrieve all indeces of an element inside the array'
Function from_dynamic_array_get_all_indeces_occurence_of(ByRef my_array,ByVal element)
  If dynamic_array_contains(my_array,element) Then
    Dim temp_array
    temp_array = Array()
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
      from_dynamic_array_get_all_indeces_occurence_of = temp_array
  Else
    Call Err.Raise(vbObjectError + 10, "Dynamic_arrays - from_dynamic_array_get_all_indeces_occurence_of", "The element "&element&" is not present")
  End If
End Function

'Function to retrieve dynamic array dimension'
Function get_dynamic_array_dimension(ByRef my_array)
  get_dynamic_array_dimension = UBound(my_array) + 1
End Function

'Funtion to write the entire array'
Function write_dynamic_array(ByRef my_array)
  Dim temp
  Dim temp_index
  temp_index = 0
  Dim array_dim
  array_dim = UBound(my_array)
  For Each temp In my_array
    If temp_index = array_dim Then
      Response.Write(temp & "<br>")
    Else
      Response.Write(temp & " - ")
    End If
    temp_index = temp_index + 1
  Next
End Function
%>
