# WCollection
 A wrapper with many useful methods for the Collection class

This class is a response to a botch job I did in a Stack Overflow answer where I provided a demonstrator as to what could be done by wrapping a Collection.  Unfortunately, the code was written in hast and had many issues about which I am now quite embarrased (https://stackoverflow.com/questions/71085037/redim-preserve-vba-alternatives/71086980?noredirect=1#comment127625279_71086980).

Hence the reason for this class.

The intent of the Wrapper is to make life easier for your typical office excel user (and also eliminate boilerplate code when coding answers to Advent Of Code problems)
Consequently, the code is functional, not clever.  Its not trying to re-implement a net environment or Net class.  It just tries to hold the hand of a programmer by doing things that might be considered useful.

The class is a self hosted factory

New instances of the class should be created using the 'Deb' method.  This method is an 'in-joke' for me as it is short for Debutante.

```
Dim myList as WCollection
Set myList = WCollection.Deb
```

As far as possible, methods return either a value or the instance of Me, to allow chaining of methods.

As of the intial upload there are 146 Rubberduck unit tests, all of which are passing.
The VBA code compiled into a twinBasic 64 bit ActiveX.dll is also provided but is currently showing some issues as 86 of the tests fail when referencing the Activex.

## Some Terminology/Organisation Stuff ##
<ul>
<li>WCollection stands for Wrapped Collection</li>
<li>An Enumerable is any VBA entity that can be enumerated using 'For ... Each'
<li>All method parameters are prefixed by ip, op or iop due to the inadequacies of ByVal and ByRef 
<ul>
<li>ip - Input parameter only. should not be passed ByRef
<li>op - Output parameter only. should not be passed ByVal.  No input value expected
<li>iop - Provides an input value which may be mutated, should not be passed ByVal
</ul>
<li>Variables global to a Class are encapsulated in one of two private UDT.  A Properties UDT with a private variable p, or a State UDT with a variable name of s.
<li>Classes with a PredecalredId are preferred to Modules.
<li>Classes tend to be written as self factories with methods that allow as much chaining as reasonable.
 
 
</ul>
