# ArrayMan
This library lets you sort the last dimension of an array of Long and ULong integers (signed and unsigned 8-byte integers), It works very fast.

The example in _Module1.vb_ demonstrates the usage of library "ArrayMan.dll", which is the only file you need to sort arrays in your projects. In the example, you can also compare the time spent for sorting with ArrayMan and the time spent with the default VB sorting algorithm.

In future, I'm planning to add support for Integer, Byte, Single and Double arrays. I can also add the ability to sort other dimensions, not only the last one.

ArrayMan.dll can only be called from 64 bit processes. 

If you will try to call it from a language other than VB.NET, the procedure expects 3 arguments: 
  1. <in> Pointer to the beginning of array of 64-bit integers to sort
  2. <out> Pointer to the beginning of array of 32-bit integers which will hold indices of the array in first argument. Its size must be at        least the number given in 3rd argument
  3. Number of items to examine in array 
  
  Procedures to call: **SortLng** for signed and 
                      **SortULng** for unsigned 64-bit ints
