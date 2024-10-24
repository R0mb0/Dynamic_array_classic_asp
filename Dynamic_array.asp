<%
'------- Variable part--------'

'Declaration variables'
Dim my_array()
Dim array_dimension
Dim array_initializated_stauts

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
get_array_dimension=array_dimension
End Function

'Funtion to get the entire array'
Function get_array()
get_array=my_array
End Function

'Function to add an element'
Function add_element(element)
If Not array_initializated_stauts Then 'If the array is not initializated'
initialize_array()
my_array(array_dimension)= element
ElseIf array_initializated_stauts and array_dimension = 0 and IsNull(my_array(array_dimension)) Then 'If the array has been initializated yet'
my_array(array_dimension)= element 
Else'If ther's no special case'
array_dimension = array_dimension + 1
Redim Preserve my_array(array_dimension)
my_array(array_dimension) = element
End If
End Function

'Function to get an element'
Function get_element(idx)
If idx <= array_dimension Then
get_element = my_array(idx)
Else
get_element = Null
End If
End Function

'Function to remove last element'
Function remove_last()
array_dimension = array_dimension - 1
Redim Preserve my_array(array_dimension)
End Function

'Function to remove an element from the array'
Function remove_element(element)
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
add_element(temp)
Next
End Function

'Funtion to check if an element is in the array'
Function contain(element)
Dim temp
For Each temp In my_array
If temp = element Then
contain = true
Exit Function
End If
Next
contain = false
End Function

'Function to retrieve the index of an element in the array'
Function index_of(element)
Dim temp_index
temp_index = 0
Dim temp
For Each temp In my_array
If temp = element Then
index_of = temp_index
Exit Function
End If
temp_index = temp_index + 1
Next
End Function

'Funtion to write the entire array'
Function write_array()
Dim temp
Dim temp_index
temp_index = 0
For Each temp In my_array
If temp_index = array_dimension Then
Response.Write(cstr(temp))
Else
Response.Write(cstr(temp) + " - ")
End If
temp_index = temp_index + 1
Next
End Function
%>