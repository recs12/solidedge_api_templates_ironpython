# -*- coding: utf-8 -*-

"""
Query in SolidEdge:
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
    selectSet.SuspendDisplay()
    selectSet.RemoveAll()

    # Query:
    # Tanks to query feature you can select elements with specified characteristics.
    # all the queries are saved in a collection  attached to the assembly -> Queries
    objQueries = asm.Queries
    print("count queries: %s" % objQueries.Count)


    # objQuery = objQueries.Item("Hardware Plated Zinc")
    # print("Query: Hardware plated zinc")
    # print("count: %s" % objQuery.MatchesCount.ToString())

    # objQuery = objQueries.Item("Hardware SS")
    # print("Query: Hardware SS")
    # print("count: %s" % objQuery.MatchesCount.ToString())

    # objQuery = objQueries.Item("fastener imperial")
    # print("Query: Hardware imperial")
    # print("count: %s" % objQuery.MatchesCount.ToString())

    # objQuery = objQueries.Item("fastener metric")
    # print("Query: Hardware metric")
    # print("count: %s" % objQuery.MatchesCount.ToString())


    # objQuery = objQueries.QuickQuery

    # Add the query here:
    zinc = objQueries.Add("Updating")
    print(zinc.MatchesCount.ToString())
    zinc.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    zinc.SearchSubassemblies = False

    # Add Criteria to above query
    zinc.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add a second criteria
    zinc.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "JDESTRX_A",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "ZINC PLATED",
    )

    # count elements
    print("Query slimane created")
    print(zinc.MatchesCount.ToString())


    # 
    # active the query here
    objQuery = objQueries.Item("Updating")

    # active the selection here
    objSelectSet = asm.SelectSet
    for occurence in objSelectSet:
        occurence.Visible = False

    # Remove query in the collection of queries
    ## objQueries.Remove("Zinc")
    ## print("Query zinc removed")

    # Updating the Screen

    # Re-enable selectset UI display.
    selectSet.ResumeDisplay()
    #  Manually refresh the selectset UI display.
    selectSet.RefreshDisplay()



def confirmation(func):
    response = raw_input("""Make a query and select ZINC PLATED elements? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes", "oui"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)
