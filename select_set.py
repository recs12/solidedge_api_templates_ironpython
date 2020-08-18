# -*- coding: utf-8 -*-

"""
SelectSet Collection Members:
==================

Public Methods
--------------
Public Method Add	            Adds an occurrence of the referenced object.
Public Method AddAll	        Enables you to select all the objects on the active sheet.(Draft)
Public Method Copy	            Places a copy of the referenced object on the clipboard.
Public Method CopyProfile
Public Method Cut	            Moves the referenced object to the system clipboard.
Public Method CutProfile
Public Method Delete	        Deletes the referenced object.
Public Method Item	            The items in a collection.
Public Method RefreshDisplay	Refreshes the select set display.
Public Method Remove	        Removes a specified object from the referenced collection.
Public Method RemoveAll	        Removes all objects from the referenced object.
Public Method ResumeDisplay	    Resumes select set display.
Public Method SuspendDisplay	Suspends select set display.


Public Properties
--------------
Public Property Application	    Returns the active application object.
Public Property Count	        Returns the number of objects in the referenced collection.
Public Property Parent	        Returns the parent object for the referenced object.
Public Property Type	        Specifies the type of the object being referenced.
"""

import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Runtime.InteropServices as SRI


def main():
    application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
    asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)

    # asm.Type =>  plate :4 , assembly : 3, partdocument: 1
    assert asm.Type == 3, "This macro only works on .asm"

    # ActiveSelectSet
    #================
    # does apply on application
    # It's the collection of elements selected in solidedge.
    # You can get this collecion just by calling it as bellow

    selectSet = application.ActiveSelectSet
    selectSet.SuspendDisplay()
    selectSet.RemoveAll() # often you want to make sure nothing is selected.

    # You can add an occurrence to the set. 
    # (Also part can be selected and add to selectSet with help of queries.)
    asm.SelectSet.Add(asm.Occurrences.item(1))
    asm.SelectSet.Add(asm.Occurrences.item(2))

    #
    print(
        "number of selected items: %s" % selectSet.Count
    )
    print(asm.selectSet[1].Name) # name of the file
    print(asm.selectSet[1].PartFileName) # path of the part
    # print(dir(asm.selectSet[1]))


    # SelectSet
    #================
    # SelectSet apply on assemblies
    #
    objSelectSet = asm.SelectSet
    for occurence in objSelectSet:
        occurence.Visible = False
    #
    # Re-enable selectset UI display.(in  the assembly tree)
    selectSet.ResumeDisplay()
    #  Manually refresh the selectset UI display.
    selectSet.RefreshDisplay()




def confirmation(func):
    response = raw_input("""Run the template (ActiveSelectSet/SelecSet)? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)
