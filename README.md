# List_GenericNLinq
## VBC.Lists.Generic and some Linq-related functions
a para-generic List-Class in VBC.  
in VBC we have the following possibilities to use a container for objects:  
a) Collection  
b) Array togehter with Redim and Redim Preserve  
c) Class holding internally and using either an Array or a Collection or both  
  
a) Collection  
Pro: 1) can hold arbitrary data-types (except for Udts coming not from tlb/axdll)  
Pro: 2) already has some useful functions: Add, Count, Item, Remove, For Each  
Pro: 3) already has some basic hashing-functionality  
Contra: 1) Memory comsumption can be rather high, because every element (even byte) is stored as a Variant (=16byte).  
Contra: 2) Data is not contained in contiguous memory   
Contra: 3) there is no possilibity for direct accessing the elements and no possibilities to access the data via pointer. You have to take the Collection-class as it is.  
  
b) Array  
Pro: 1) can hold every Datatype, even non opublic udts (User-defined-Types, Types)  
Pro: 2) for every Datatype basically as much memory will be used like the data itself, no data overhead  
Contra: 1) for every Array you need it's own function for dealing with the data. Moreover all the time some mor or less important functions are anywhere in you code which maybe call Redim wo Preserve.  
  
c) Class  
if you act smart you can use best of both worlds from Collection and Array and eliminate the disadvantages:  
For sequential data use an Array, for hashing-funktionality use a collection.  
  
Why para-generic?  
a generic class will be typified during design time. During compile time the compiler creates it's own class for every datatyp in use.  
there is nothing like that in VBC. If there wasn't the Variant-datatype, you had to implement it's own class for every datatype by hand by yourself.   
Thus the same code will be copied over and over again. a possible extension to the class has to be copied again for each list-class.  
  
But yes we can, determine the datatype of your objects during run-time!  
during design time of the list we just need to deliver a variable of type VBVartype, und in the list-class and from now on the datatype ca nbe determoined durring run-time.   
As datapointer (resp. as Array-variable) there is a variant applied.  

Finally the list became some functions to mimic a behaviour very much like Linq in .net.  

![GenericListLinq Image](Resources/GenericListLinq.png "GenericListLinq Image") 