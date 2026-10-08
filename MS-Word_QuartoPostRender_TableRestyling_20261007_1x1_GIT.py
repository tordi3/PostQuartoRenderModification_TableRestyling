#!/usr/bin/env python
# coding: utf-8

# # Preamble and notes
# 
# <!-- 
# 
# This script is run using the below programming language
# 
# Programming Language = Python - Version 3.13.7 - 32-bit
# 
# -->

# # Import necessary modules 

# In[2]:


import os #needed for navigating to files within the operating system

import io
import sys

import docx
from docx import Document
from lxml import etree
import zipfile

print("Python version:")
print(sys.version)

print("docx version")
print(docx.__version__)


# # Determine current working directory
# 

# In[3]:


current_directory = os.getcwd()
print("Current directory is below")
print(current_directory)
print("Current directory is above")


# # Set variables for relevant directories
# Set path to the target working directory to be able to access source information file

# In[4]:


# Only works outside of Jupyter note book
script_path = os.path.dirname(os.path.abspath(__file__))

#Works in Jupyter notebook, but not outside of Jupyter notebook
#script_path =  globals()['_dh'][0]


# ## Set path to source information file
# 

# In[5]:


#UserInput_DataFileEntry = input("Input name of docx data file, which needs specific tables updated, in here: ")
UserInput_DataFileEntry = "example_testsample1.docx"

UserInput_QuartoTblGrphcsKey = input("User please input the Quarto Table Graphics Key here (ex: tbl (The standard prefix used by quarto), tab (custom prefix), stab (custom prefix)): ")

#DataFileEntry = "ExtractContent.docx"



DataFileNameModificiation = UserInput_DataFileEntry.replace(".docx","")
#DataFileNameModificiation_1 = DataFileEntry.replace("_OrganizedDataOutput","")
DataFileNameModificiation_1 = f"{DataFileNameModificiation}_TablesUpdated"
data_file = os.path.join(script_path, UserInput_DataFileEntry)


# ## Set path to output data file - csv

# In[6]:


OutputFileExtension = ".docx"

#OutputFileName_1 = f"{OutputFileName}_OrganizedDataOutput{OutputFileExtension}"
OutputFileName = f"{DataFileNameModificiation_1}{OutputFileExtension}"


OutputFile = os.path.join(script_path, OutputFileName)


print(OutputFile)



# # Main Document (document.xml) Processing

# In[7]:


# Namespace Format Definition
ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# MS Word XML namespace prefix definition for the "w" namespace (wordprocessingml) in the document.xml file.
namespaceprefix_1x1 = f"{{{ooXMLns['w']}}}"

# Source docx file reference
document = Document(data_file)

# Source docx file open variable definition. 
docxZip = zipfile.ZipFile(data_file)

# Docx file main document with bookmarked tables reading definition
# Within the docx zip "word" directory, go to the "document.xml" file and open it for reading. 
# The document.xml file contains all the table metadata (table id, style, etc) in the docx file.
documentXML = docxZip.read('word/document.xml')

# Define the XML tree structure for the document.xml file
# (documentXML) is considered the root of the XML tree structure for the document.xml file.
et_1x1 = etree.XML(documentXML) 

# When the user want to change table style for a specific bookmarked table (like support tables (probably stbl-) or main tables(probably tbl-))
#Str_QuartoTblType = "tbl" + "-"
Str_QuartoTblType = UserInput_QuartoTblGrphcsKey + "-"
print(f"The user selected the table type >>> {Str_QuartoTblType}")

# Provide a list of all the quarto bookmarked tables being changed
print("This is the list of all the quarto bookmarked tables in the document:")
Path_QuartoBKMdTbl = f'(//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])'
List_QuartoBKMdTbl = et_1x1.xpath(Path_QuartoBKMdTbl,namespaces=ooXMLns) 
print(f'>>> {List_QuartoBKMdTbl}')

QtyValue_QuartoBKMdTbl = len(List_QuartoBKMdTbl)
print(f'The number of quarto bookmarked tables in the document is: {QtyValue_QuartoBKMdTbl}')

print("This is an assumption based off the first found table, with the selected Quarto Table Graphics Key.")
print("The tables with selected Quarto Table Graphics Key have the following table style attribute value of: ")
ListPosNum_FrstItem_TarBKMdTblIdDef_1x1 = 1
Path_FrstItem_TarBKMdTblIdDef_1x1 = f'((//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])[{ListPosNum_FrstItem_TarBKMdTblIdDef_1x1}])/following-sibling::w:tbl/descendant::w:tblStyle/@w:val'
List_FrstItem_TarBKMdTblIdDef_1x1 = et_1x1.xpath(Path_FrstItem_TarBKMdTblIdDef_1x1,namespaces=ooXMLns) 
print(f'>>> {List_FrstItem_TarBKMdTblIdDef_1x1[0]}')


UserInput_NewAttribute_1x1 = input("User, please input the 'w:styleId' value of the new table style you wish to set the selected tables. "
"The value of the intended 'w:styleId' can be found in the 'styles.xml' of your '.docx' "
"\nThese need to be tables you have already defined in your document style reference (the Quarto reference-doc). "
"\n('TableGrid' is a MS word default, and will attempt to be selected if nothing is entered): ")
if UserInput_NewAttribute_1x1 == "":
    UserInput_NewAttribute_1x1 = "TableGrid"
    print(f'User selected the default table style >>> {UserInput_NewAttribute_1x1}.')
elif not UserInput_NewAttribute_1x1 == "":
    UserInput_NewAttribute_1x1
    print(f'User selected the table style >>> {UserInput_NewAttribute_1x1}.')




for index_TargetBMKdTbl in range(1, QtyValue_QuartoBKMdTbl + 1):
    print("The current table being changed has the following identification number:")
    Path_current_tbl_id = f'((//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])[{index_TargetBMKdTbl}]//@w:id)'  # Replace with the actual tbl-id
    List_current_tbl_id = et_1x1.xpath(Path_current_tbl_id,namespaces=ooXMLns) 
    print(f'>>> {List_current_tbl_id[0]}')

    # Need to select the entire node for changing the table style attribute value.
    # Verify the current table style for the target.
    print("The current table being changed has the table style attribute value of:")
    Path_TargetBookmarkedTableIdDefinition_1x1 = f'((//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])[{index_TargetBMKdTbl}])/following-sibling::w:tbl/descendant::w:tblStyle/@w:val'
    List_TargetBookmarkedTable_1x1 = et_1x1.xpath(Path_TargetBookmarkedTableIdDefinition_1x1,namespaces=ooXMLns) 
    print(f'>>> {List_TargetBookmarkedTable_1x1[0]}')

    # Go the target Bookmarked Table Style Node 
    print("Going to the target Bookmarked Table Style Node...")
    #Item_TargetBookmarkedTableIdDefinition_1x1 = f'((//w:bookmarkStart[contains(@w:name,"tbl-")])[1])/following-sibling::w:tbl/descendant::w:tblStyle/@w:val'
    Path_TargetBookmarkedTableIdDefinition_1x1 = f'((//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])[{index_TargetBMKdTbl}])/following-sibling::w:tbl/descendant::w:tblStyle'
    List_TargetBookmarkedTable_1x2 = et_1x1.xpath(Path_TargetBookmarkedTableIdDefinition_1x1,namespaces=ooXMLns) 
    print(f'>>> {List_TargetBookmarkedTable_1x2}')

    # Select the table style attribute value for the target bookmarked table in the document.
    # Define the selection of the target bookmarked table style attribute value for the target bookmarked table in the document.
    print("The target attribute title (key) is:")
    target_AttributeTitle_1x1 = 'val'
    target_Attribute_1x1 = f"{namespaceprefix_1x1}{target_AttributeTitle_1x1}" # w can't be used as the namespace string, it needs to be called directly
    print(f'>>> {target_Attribute_1x1}')

    # At the target bookmarked table style node, change its value from the current value "Table" (which is the default table style in Word), 
    ## to the user defined style (ex: TableGrid).
    # Notification of attribute change occuring.
    print("Changing the table style for the target bookmarked table...")
    #List_TargetBookmarkedTable_1x1[0].get(namespacestring) # Not working, was intended to for viewing the attribute before changing it.
    List_TargetBookmarkedTable_1x2[0].set(target_Attribute_1x1, UserInput_NewAttribute_1x1)

    print("The new table style attribute value is:")
    # Verify the change of the table style for the target bookmarked table in the document.
    Path_TargetBookmarkedTableIdDefinition_1x3 = f'((//w:bookmarkStart[contains(@w:name,"{Str_QuartoTblType}")])[{index_TargetBMKdTbl}])/following-sibling::w:tbl/descendant::w:tblStyle/@w:val'
    List_TargetBookmarkedTable_1x3 = et_1x1.xpath(Path_TargetBookmarkedTableIdDefinition_1x3,namespaces=ooXMLns) 
    print(f'>>> {List_TargetBookmarkedTable_1x3[0]}')

# Saving the modified XML within a new docx file.
# Defining the modified XML tree structure for the document.xml file

# et_1x1 contains the updated root.
et_2x1 = etree.ElementTree(et_1x1)
modified_xml = etree.tostring(et_2x1, encoding='utf-8', xml_declaration=True)
print(etree.tostring(et_2x1, pretty_print=True, encoding='utf-8', xml_declaration=True))

# Defining temporary zip file for saving the modified XML within a new docx file.
temp_file = data_file + '.temp'

# Re-open the source zip to read, and open a new temporary zip to write out changes
with zipfile.ZipFile(data_file, 'r') as zin, zipfile.ZipFile(temp_file, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'word/document.xml':
            # Write your freshly updated XML bytes instead of the original file
            zout.writestr(item, modified_xml)
        else:
            # Copy all other pristine files (styles, images, numbering, etc.) exactly as-is
            zout.writestr(item, zin.read(item.filename))

# Safely close open zip references by stepping out of context blocks, then replace the file
os.replace(temp_file, OutputFile)
print(f"Successfully saved all table style modifications back to: {OutputFile}")

