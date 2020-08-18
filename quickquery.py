# -*- coding: utf-8 -*-

"""
QuickQuery in SolidEdge:
==================
"""

import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Runtime.InteropServices as SRI
from System import Console
import SolidEdgeAssembly as SEAssembly


def main():
    application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
    asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)

    # asm.Type =>  plate :4 , assembly : 3, partdocument: 1
    assert asm.Type == 3, "This macro only works on .asm"

    # ActiveSelectSet
    # It's the collection of elements selected in solidedge.
    # You can get this collecion just by calling it as bellow
    selectSet = application.ActiveSelectSet

    # Query:
    # Tanks to query feature you can select elements with specified characteristics.
    # all the queries are saved in a collection  attached to the assembly -> Queries
    objQueries = asm.Queries
    print("Count queries: %s" % objQueries.Count)


    # Quickquery here:
    # quick = objQueries.Add("Updating")
    quick = objQueries.QuickQuery
    quick.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    quick.SearchSubassemblies = False

    # Add Criteria to above query
    quick.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )

    # count elements
    print("quick query created")


    # active the query here
    print(quick.MatchesCount.ToString())

    # active the selection here
    objSelectSet = asm.SelectSet
    for occurence in objSelectSet:
        occurence.Visible = False


def confirmation(func):
    response = raw_input("""Make a query and select ZINC PLATED elements? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes", "oui"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)
