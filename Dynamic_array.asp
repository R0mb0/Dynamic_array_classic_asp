<%@LANGUAGE="VBSCRIPT"%>
<%

'------- Variable part--------'

'Declaration variables'
Dim my_array()
Dim array_dimension
Dim array_initializated_stauts

'Initialize variables'
array_initializated_stauts = false
array_dimension = Null

'-------- Functions part--------'

'Funtion to get array status'
Function is_array_initializated()
is_array_initializated = array_initializated_stauts
End Function

'Function to initialize the array'
Function initialize_array()
array_dimension = 0
ReDim my_array(array_dimension)
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
If Not array_initializated_stauts 'If the array is not initializated'
initialize_array()
my_array(array_dimension)= element
ElseIf array_initializated_stauts && array_dimension=0 && IsNull(my_array(array_dimension))'If the array has been initializated yet'
my_array(array_dimension)= element 
Else'If ther's no special case'
array_dimension = array_dimension + 1
Redim Preserve my_array(array_dimension)
my_array(array_dimension) = element
End If
End Function

'Function to get an element'
Function get_element(indice)
If indice <= array_dimension
get_element = my_array(indice)
Else
get_element = Null
End Function

'Function to remove last element'
Function remove_last()

End Function

%>
