# Dynamic array in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/5fb4b710ca5c4dfd88a883944af2dac3)](https://app.codacy.com/gh/R0mb0/Dynamic_array_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Dynamic_array_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Dynamic_array_classic_asp)
[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `Dynamic_array.asp`'s avaible Functions

- **Initialize the array** -> `initialize_array()`
- **Check if the array is initializated** -> `is_array_initializated()`
- **Add an element into the array** -> `add_element_to_array(element)`
- **Get element from index** -> `get_element_from_array(idx)`
- **Remove last element from array** -> `remove_last_element_from_array()`
- **Remove an element from array** -> `remove_all_occurences_from_array(element)`
- **Remove the first element occurence from array** -> `remove_first_occurence_from_array(element)`
- **Remove elements from indices** -> `remove_these_elements_from_array(indexes_array)`
- **Reset the array** -> `initialize_array()`
- **Check if an element is in the array** -> `array_contains(element)`
- **Retrieve the index of an element in the array** -> `from_array_get_index_of(element)`
- **Retrieve all indeces of an element in the array (return an array)** -> `from_array_get_all_indeces_occurence_of(element)`
- **Retrieve the entire array** -> `get_array()`
- **Retrieve the array dimension** -> `get_array_dimension()`
- **Write the entire array** -> `write_array()`

## How to use: 

> From `Test1.asp`

1. Initialize the array and check it's status
   ```
   initialize_array()
   Response.Write("Array status: ")
   Response.Write(is_array_initializated() & "<br>")
   ```
2. Use the functions to manage the array
   ```
   add_element_to_array("A")
   add_element_to_array("B")
   add_element_to_array("C")
   add_element_to_array("D")
   Response.Write("Elements inside: ")
   write_array()
   ```
