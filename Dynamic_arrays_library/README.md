# Dynamic arrays in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/5512538b7d044e76a4c3052dd5032b80)](https://app.codacy.com/gh/R0mb0/Dynamic_arrays_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Dynamic_arrays_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Dynamic_arrays_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `Dynamic_arrays.asp`'s avaible Functions

- **Initialize a dynamic array** -> `get_initializated_dynamic_array()`
- **Add an element into a dynamic array** -> `add_element_to_dynamic_array(my_array,element)`
- **Get element from index** -> `get_element_from_dynamic_array(my_array,idx)`
- **Remove last element from a dynamic array** -> `remove_last_element_from_dynamic_array(my_array)`
- **Remove all element occurences from a dyamic array** -> `remove_all_occurences_from_dynamic_array(my_array,element)`
- **Remove first element occurence from a dynamic array** -> `remove_first_occurence_from_dynamic_array(my_array,element)`
- **Remove element from index** -> `remove_this_elements_from_dynamic_array(my_array,idx)`
- **Remove elements from indices** -> `remove_these_elements_from_dynamic_array(my_array,indices_array)`
- **Reset a dynamic array** -> `get_initializated_dynamic_array()`
- **Check if an element is in the dynamic array** -> `dynamic_array_contains(my_array,element)`
- **Retrieve the first index of an element inside a dynamic array** -> `from_dynamic_array_get_first_index_occurence_of(my_array,element)`
- **Retrieve all indeces of an element inside the dynamic array (return an array)** -> `from_dynamic_array_get_all_indeces_occurence_of(my_array,element)`
- **Retrieve the dynamic array dimension** -> `get_dynamic_array_dimension(my_array)`
- **Write an entire dynamic array** -> `write_dynamic_array(my_array)`

## How to use: 

> From `Test.asp`

1. Create array and initialize it
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="Dynamic_arrays.asp"-->
   <%
   Dim test_array
   test_array = Array()
   test_array = get_initializated_dynamic_array()
   ```
2. Pass the array to functions for manage
   ```
   add_element_to_dynamic_array test_array,"A"
   add_element_to_dynamic_array test_array,"B"
   add_element_to_dynamic_array test_array,"C"
   write_dynamic_array(test_array)
   %>
   ```
