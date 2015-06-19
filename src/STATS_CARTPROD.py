#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2014
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

from __future__ import with_statement

"""STATS CARTPROD extension command"""

__author__ =  'IBM SPSS, JKP'
__version__=  '1.0.3'

# history
# 28-feb-2014 original version
# 24-sep-2014 fix problem with upper case variable names
# 21-nov-2014 TO and ALL support

helptext = """STATS CARTPROD INPUT2=datasetname
    VAR1=variable list VAR2=variable list
/SAVE OUTFILE="filespec" DSNAME=dataset name
/HELP

Example:
STATS INPUT2=right
    VAR1=x1 x2 VAR2=y1 y2
/SAVE OUTFILE=cartprod.sav DSNAME=cart.

This command computes the cartesian product of the VAR1 variables
with the VAR2 variables and produces a new data file containin the result.

The active file must have a dataset name in order to run this command.

All keywords and subcommands are required except INPUT2 and DSNAME.

The active dataset is the left term in the cartesian product.
INPUT2 specifies the dataset containing the right hand variables to be used.
If INPUT2 is omitted, all variables are taken from the active dataset.

VAR1 and VAR2 specify the variable(s) whose cartesian product is  computed.
The two lists cannot have any names in common.

OUTFILE specifies a file name for the output.  It cannot be a dataset name.
The output file cannot be the active file.  It will be open after the
command completes.

DSNAME specifies a dataset name for the output data file.  If no name is
specified, a random name will be assigned.

/HELP displays this help and does nothing else.
"""

import spss, spssaux
from extension import Template, Syntax, processcmd
import  random, tempfile

def docart(var1, var2, outfile, dsname=None, input2=None):
    """Create dataset containing cartesian product of selected variables"""

    activedsname, casecount1, casecount2 =  dscheck(input2, var1, var2)
    
    # To compute the cartesian product
    # 1. Replicate each case in dataset 1 by the number of cases in ds 2 into a new dataset
    # 2. Replicate each case in dataset 2 by the number of cases in ds 1
    #    adding a 1:Card(ds1) sequence number
    # 3. Sort ds2 by the sequence number
    # 4. ADD VARIABLES from ds2 to ds1 dropping the sequence number
    
    var1str = " ".join(var1)
    var2str = " ".join(var2)
    randomvname = "V" + str(random.uniform(.1, 1))
    temp = tempfile.mktemp() + ".sav"
    randomdnametemp = "D" + str(random.uniform(.1, 1))
    if dsname is None:
        randomdnamemain = "D" + str(random.uniform(.1, 1))
    else:
        randomdnamemain = dsname
    
    # 1.
    # If both variable sets come from primary input, create both files on one data pass.
    # Otherwise, pass both separately.
    if input2:
        cmd ="""LOOP #i = 1 TO %(casecount2)s.
        XSAVE OUTFILE="%(outfile)s" /KEEP=%(var1str)s.
        END LOOP.""" % locals()
    else:
        cmd = """TEMPORARY.
        LOOP %(randomvname)s = 1 TO %(casecount1)s.
        XSAVE OUTFILE="%(outfile)s" /KEEP=%(var1str)s.
        XSAVE OUTFILE="%(temp)s" /KEEP %(randomvname)s %(var2str)s.
        END LOOP.""" % locals()
    spss.Submit(cmd)
    
    # 2.  This step accomplished in #1 if all inputs from one dataset.
    if input2:
        cmd = """DATASET ACTIVATE %(input2)s.""" % locals()
        spss.Submit(cmd)
        cmd = """TEMPORARY.
        LOOP %(randomvname)s = 1 TO %(casecount1)s.
        XSAVE OUTFILE="%(temp)s" /KEEP %(randomvname)s %(var2str)s.
        END LOOP.""" % locals()
        spss.Submit(cmd)
    
    # 3.
    cmd = """GET FILE "%(temp)s".
DATASET NAME %(randomdnametemp)s.
SORT CASES BY %(randomvname)s.""" % locals()
    spss.Submit(cmd)
    
    #4.
    cmd = """GET FILE="%(outfile)s".
DATASET NAME %(randomdnamemain)s.
MATCH FILES /FILE=*
/FILE = %(randomdnametemp)s
/DROP= %(randomvname)s.
EXECUTE.
DATASET CLOSE %(randomdnametemp)s.
ERASE FILE="%(temp)s".
DATASET ACTIVATE %(randomdnamemain)s WINDOW=FRONT.""" % locals()
    spss.Submit(cmd)
    

def dscheck(input2, var1, var2):
    """Check dataset and variable conditions and return name of active ds and case counts
    
    input2 is the name of the second dataset or None
    var1 and var2 are the variable lists involved
    outputds is the name of the output dataset to be created"""
    
    # var1 names are already validated
    # var2 validation uses caseless option for the VariableDict in order
    # to accommodate CDB limitations
    
    activedsname = spss.ActiveDataset()
    var1 = set([v.lower() for v in var1])
    var2 = set([v.lower() for v in var2])
    if activedsname == "*":
        raise ValueError(_("""The active dataset must have a dataset name in order to run this command"""))
    if var1.intersection(var2):
        raise ValueError(_("""The two variable lists cannot have any names in common"""))
    if input2 is not None:
        input2 = input2.lower()

    casecount1 = spss.GetCaseCount()
    if casecount1 < 0:
        spss.Submit("EXECUTE")
        casecount1 = spss.GetCaseCount()
        
    if input2:
        spss.Submit("DATASET ACTIVATE %s" % input2)   # fails if does not exist
        try:
            varnames2 = set(item.lower() for item in spssaux.VariableDict(caseless=True).variables if item in var2)
        except(TypeError):
            raise TypeError(_("""This command requires a newer version of the spssaux.py module"""))
        if len(varnames2) != len(var2):
            raise ValueError(_("""The right side variable list contains undefined variables:\n%s""")\
                % ", ".join(var2 - varnames2))
        casecount2 = spss.GetCaseCount()
        if casecount2 < 0:
            spss.Submit("EXECUTE")
            casecount2 = spss.GetCaseCount()        
        spss.Submit("DATASET ACTIVATE %s" % activedsname)
        diff = var2 - varnames2

    else:
        try:
            diff = var2 - set(item.lower() for item in spssaux.VariableDict(caseless=True).variables)
        except(TypeError):
            raise TypeError(_("""This command requires a newer version of the spssaux.py module"""))        
        casecount2 = casecount1
    if diff:
        raise ValueError(_("""The second set of input variables contains undefined variables:\n%s""")\
            % ", ".join(diff))
    return activedsname, casecount1, casecount2
    
def Run(args):
    """Execute the STATS CARTPROD extension command"""

    args = args[args.keys()[0]]
    # debugging
    # makes debug apply only to the current thread
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
        #import thread
        #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
        ## for V19 use
        ##    ###SpssClient._heartBeat(False)
    #except:
        #pass
    oobj = Syntax([

        Template("INPUT2", subc="",  ktype="literal", var="input2"),
        Template("VAR1", subc="", ktype="existingvarlist", var="var1", islist=True),
        Template("VAR2", subc="", ktype="varname", var="var2", islist=True),
        
        Template("OUTFILE", subc="SAVE",  ktype="literal", var="outfile"),
        Template("DSNAME", subc="SAVE", ktype="varname", var="dsname"),
        Template("HELP", subc="", ktype="bool")])
    
    #enable localization
    global _
    try:
        _("---")
    except:
        def _(msg):
            return msg
    # A HELP subcommand overrides all else
    if args.has_key("HELP"):
        #print helptext
        helper()
    else:
        processcmd(oobj, args, docart, vardict=spssaux.VariableDict())

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print("Help file not found:" + helpspec)
        
class NonProcPivotTable(object):
    """Accumulate an object that can be turned into a basic pivot table once a procedure state can be established"""
    
    def __init__(self, omssubtype, outlinetitle="", tabletitle="", caption="", rowdim="", coldim="", columnlabels=[],
                 procname="Messages"):
        """omssubtype is the OMS table subtype.
        caption is the table caption.
        tabletitle is the table title.
        columnlabels is a sequence of column labels.
        If columnlabels is empty, this is treated as a one-column table, and the rowlabels are used as the values with
        the label column hidden
        
        procname is the procedure name.  It must not be translated."""
        
        attributesFromDict(locals())
        self.rowlabels = []
        self.columnvalues = []
        self.rowcount = 0

    def addrow(self, rowlabel=None, cvalues=None):
        """Append a row labelled rowlabel to the table and set value(s) from cvalues.
        
        rowlabel is a label for the stub.
        cvalues is a sequence of values with the same number of values are there are columns in the table."""

        if cvalues is None:
            cvalues = []
        self.rowcount += 1
        if rowlabel is None:
            self.rowlabels.append(str(self.rowcount))
        else:
            self.rowlabels.append(rowlabel)
        self.columnvalues.extend(cvalues)
        
    def generate(self):
        """Produce the table assuming that a procedure state is now in effect if it has any rows."""
        
        privateproc = False
        if self.rowcount > 0:
            try:
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            except:
                StartProcedure(_("Adjust Widths"), self.procname)
                privateproc = True
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            if self.caption:
                table.Caption(self.caption)
            if self.columnlabels != []:
                table.SimplePivotTable(self.rowdim, self.rowlabels, self.coldim, self.columnlabels, self.columnvalues)
            else:
                table.Append(spss.Dimension.Place.row,"rowdim",hideName=True,hideLabels=True)
                table.Append(spss.Dimension.Place.column,"coldim",hideName=True,hideLabels=True)
                colcat = spss.CellText.String("Message")
                for r in self.rowlabels:
                    cellr = spss.CellText.String(r)
                    table[(cellr, colcat)] = cellr
            if privateproc:
                spss.EndProcedure()
                
def attributesFromDict(d):
    """build self attributes from a dictionary d."""
    self = d.pop('self')
    for name, value in d.iteritems():
        setattr(self, name, value)

def StartProcedure(procname, omsid):
    """Start a procedure
    
    procname is the name that will appear in the Viewer outline.  It may be translated
    omsid is the OMS procedure identifier and should not be translated.
    
    Statistics versions prior to 19 support only a single term used for both purposes.
    For those versions, the omsid will be use for the procedure name.
    
    While the spss.StartProcedure function accepts the one argument, this function
    requires both."""
    
    try:
        spss.StartProcedure(procname, omsid)
    except TypeError:  #older version
        spss.StartProcedure(omsid)