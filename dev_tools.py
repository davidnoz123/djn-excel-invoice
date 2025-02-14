r"""


C:\analytics\projects\git\lexi\demos\venv\Scripts\python.exe

import runpy ; temp = runpy._run_module_as_main("dev_tools")


"""

import sys
import os
import threading
import re
# import time
# import json
# 
# 
# import importlib
# import collections
# import datetime
# import signal
# import functools
# import uuid
# import zipfile
# import xml.etree.ElementTree as ET
# import random
# import zlib
# import base64
# import enum
# import io
# import traceback

file__fullPath = os.path.abspath(__file__)
file__baseName = os.path.basename(file__fullPath)
file__parentDr = os.path.dirname(file__fullPath)
file__fileSysD = (lambda a:lambda v:a(a, v, v))(lambda s, v, x:x if os.path.isdir(x) else (_ for _ in ()).throw(Exception(f"Argument not a directory:'{v}'")) if x==os.path.dirname(x) else s(s, v, os.path.dirname(x)))(file__parentDr)



class classproperty(object):
    def __init__(self, f):
        self.__qualname__ = f.__qualname__
        self.f = f
    def __get__(self, obj, owner):
        return self.f(owner)    
    def __call__(self, f):
        def _f(*args, **kwargs):
            return f(*args, **kwargs)
        return _f

class ExcelVBA:

    vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document = 1, 2, 3, 100
    
    vbext_ct_ComponentTypes = {vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document}
    
    # https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#type
    vbext_ct_ComponentTypesx = {vbext_ct_StdModule : ("StdModule", "bas"), vbext_ct_ClassModule : ("ClassModule", "cls"), vbext_ct_MSForm : ("MSForm", "frm"), vbext_ct_Document : ("Document", "cls")}    

    @classmethod
    def validate_tlx_base(cls, tlx):        
        ret = (isinstance(tlx, tuple) and len(tlx) == 2)
        return ret
        
    @classmethod
    def validate_tl(cls, tl):
        ret = cls.validate_tlx_base(tl) and isinstance(tl[0], int) and isinstance(tl[1], int) and tl[0] > 0 and tl[1] > 0
        return ret
        
    @classmethod
    def validate_tlbr(cls, tlbr):
        ret = cls.validate_tlx_base(tlbr) and cls.validate_tl(tlbr[0]) and cls.validate_tl(tlbr[1])
        return ret
        
    @classmethod
    def validate_tl_or_tlbr(cls, tlx):
        ret = cls.validate_tlx_base(tlx) and ((isinstance(tlx[0], int) and isinstance(tlx[1], int) and tlx[0] > 0 and tlx[1] > 0) or (cls.validate_tl(tlx[0]) and cls.validate_tl(tlx[1])))
        return ret    

    @classmethod
    def excel_application_create(cls):
        return com_api.win32com_client_DispatchEx("Excel.Application")

    @classmethod
    def excel_application_release(cls, excel, force=False):
        try:
            com_api.win32com_client_DispatchEx_release(excel, hwnd = excel.Hwnd, force=force)
        except:
            pass

    @classmethod        
    def safe_release_excel(cls):
        if "excel_api_have_excel_application" in globals():
            tmp = globals()["excel_api_have_excel_application"]        
            try:
                com_api.win32com_client_DispatchEx_release(tmp, hwnd = tmp.Hwnd)
            finally:
                del globals()["excel_api_have_excel_application"]        
                
    @classmethod        
    def get_excel_threadsafe_x(cls):
        ret = None
        for attempt in range(2):
            try:    
                ret = cls.get_excel_x()
                break
            except BaseException as e:
                if attempt == 0 and hasattr(e, 'args') and e.args == (-2147221008, 'CoInitialize has not been called.', None, None):
                    CoInitialize()
                else:
                    raise
        return ret
    
    @classmethod         
    def get_excel_x(cls):
        # https://timgolden.me.uk/python/win32_how_do_i/attach-to-a-com-instance.html
        import win32com
        import win32com.client        
        xl = win32com.client.GetActiveObject("Excel.Application")
        try:
            xl_hWnd = xl.hWnd
        except AttributeError as e:
            if str(e) == 'Excel.Application.hWnd':
                print(f"{e.__class__}:{e}:NOTE:This error happens when Excel is in input mode i.e., when a cell is selected and taking input from the keyboard. Check if this is the case and, if so, try again.")
            raise    
        return xl                  
    
    @classproperty
    def excel(cls):
        if "excel_api_have_excel_application" not in globals():
            #globals()["excel_api_have_excel_application"] = win32com.client.gencache.EnsureDispatch('Excel.Application')
            globals()["excel_api_have_excel_application"] = cls.excel_application_create() # http://timgolden.me.uk/python/win32_how_do_i/start-a-new-com-instance.html               
        ret = globals()["excel_api_have_excel_application"]
        return ret

    win32com_client_constants = None

    class constants:
            xlValidateList      = 3 # https://learn.microsoft.com/en-us/office/vba/api/excel.xldvtype
            xlValidAlertStop    = 1 # https://learn.microsoft.com/en-us/office/vba/api/excel.xldvalertstyle
            xlBetween           = 1 # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlformatconditionoperator?view=excel-pia
            xlSrcRange          = 1 # https://learn.microsoft.com/en-us/office/vba/api/excel.xllistobjectsourcetype
            xlYes               = 1 # https://learn.microsoft.com/en-us/office/vba/api/excel.xlyesnoguess
            
            ErrDiv0  = -2146826281
            ErrNA    = -2146826246
            ErrName  = -2146826259
            ErrNull  = -2146826288
            ErrNum   = -2146826252
            ErrRef   = -2146826265
            ErrValue = -2146826273
    
    #@classproperty
    #def constants(cls):
    #    if cls.win32com_client_constants is None:
    #        pywin_import()
    #        cls.win32com_client_constants = win32com.client.constants
    #    return cls.win32com_client_constants
            
    @classmethod
    def Excel(cls):
        return cls.excel        
    
    @classmethod    
    def WsExists(cls, wb, sName):
        try:
            ws = wb.Sheets(sName)
            ret = True
        except:
            ret = False
        return ret      

    @classmethod
    def PutData(cls, top_left, a, save_changes=False):
        if len(a) > 0:
            ws = top_left.Parent
            tl = ws.Cells(top_left.Row , top_left.Column)
            br = ws.Cells(top_left.Row + len(a) - 1, top_left.Column + len(a[0]) - 1)
            ws.Range(tl, br).Value = a
            if save_changes: ws.Parent.Save()   

    @classmethod
    def ListObjectsAdd(cls, wb, sWorksheetName, tlbr, name=None, style_str=None):
        ws = wb.Worksheets(sWorksheetName)                
        (row_beg, col_beg), (row_end, col_end) = tlbr
        r = ws.Range(ws.Cells(row_beg, col_beg), ws.Cells(row_end, col_end))
        #https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.add 
        t = ws.ListObjects.Add(cls.constants.xlSrcRange, r, None, cls.constants.xlYes)
        if name is not None: t.Name = name
        t.TableStyle = "TableStyleLight1"
        return t
            
    @classmethod
    def FreezeTopRow(cls, wb, sWorksheetName, offset=0):
        app = wb.Parent        
        ws = wb.Worksheets(sWorksheetName)    
        try:
            rSelection = app.Selection
        except BaseException as e:
            rSelection = None
            #print(e)
            pass

        bScreenUpdating = app.ScreenUpdating

        try:
            app.ScreenUpdating = False
            wb.Activate()
            ws.Activate()
            ws.Cells(1 + offset, 1).Select()    
            try:
                app.ActiveWindow.FreezePanes = False
            except BaseException as e:
                raise

            app.ActiveWindow.SplitColumn = 0
            app.ActiveWindow.SplitRow = max(1, offset - 1)
            
            try:
                app.ActiveWindow.FreezePanes = True
            except BaseException as e:
                raise            
            
            if rSelection is not None:
                try:
                    rSelection.Parent.Activate
                    rSelection.Select
                except BaseException as e:
                    #print(e)
                    pass
        finally:
            app.ScreenUpdating = bScreenUpdating

    @classmethod
    def ClearAllNames(cls, ws):
        while ws.Names.Count > 0:
            ws.Names(1).Delete()

    @classmethod          
    def ClearAllShapes(cls, ws):
        ws.Pictures().Delete()
        while ws.Shapes.Count > 0:
            ws.Shapes(1).Delete()
            
    @classmethod          
    def ClearAllListObjects(cls, ws):
        while ws.ListObjects.Count > 0:
            ws.ListObjects(1).Delete()

    @classmethod
    def SafeAddWorksheet(cls, wb, sWorksheetName):
        try:
            rSelection = wb.Parent.Selection
        except:
            pass
        bScreenUpdating = wb.Parent.ScreenUpdating
        wb.Parent.ScreenUpdating = False
        ret = wb.Sheets.Add()
        wb.Parent.ScreenUpdating = bScreenUpdating
        ret.Name = sWorksheetName
        try:
            rSelection.Parent.Activate
            rSelection.Select
        except:
            pass
        return ret            

    @classmethod
    def SafeGetWorksheet(cls, wb, sWorksheetName, bEnsureEmpty=False):
        try:
            ret = wb.Sheets(sWorksheetName)
        except BaseException as e:
            ret = cls.SafeAddWorksheet(wb, sWorksheetName)            
        if bEnsureEmpty:
            cls.ClearAllListObjects(ret)            
            ret.Cells.Delete()
            cls.ClearAllNames(ret)
            cls.ClearAllShapes(ret)    
        return ret
        
    @classmethod
    def SafeGetWorkbook_SafeSaveAs(cls, wb, sPath, FileFormat, excel=None): 
        if excel is None: excel = cls.Excel()
        Application = excel     
        
        bDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        try:
            wb.SaveAs(sPath, FileFormat=FileFormat)
        finally:
            Application.DisplayAlerts = bDisplayAlerts

    @classmethod
    def SafeGetWorkbook(cls, sPath, sTemplate="", bEnsureEmpty=False, bUpdateLinks=True, excel=None):
        if excel is None: excel = cls.Excel()
        Application = excel        
        
        sPath = os.path.abspath(sPath)
        
        wb = None

        if sTemplate != "" and bEnsureEmpty:
          raise Exception("Utils.SafeGetWorkbook:Inconsistent arguments:sTemplate <> And bEnsureEmpty")
          
        sWorkbook = os.path.basename(sPath)
        
        bReopenTemplate = False
        bScreenUpdating = Application.ScreenUpdating

        wbActive = Application.ActiveWorkbook

        bFileAlreadyOpen = sWorkbook.lower() in set([wb.Name.lower() for wb in Application.Workbooks])
        
        if bFileAlreadyOpen:
            wb = Application.Workbooks[sWorkbook]
            if bEnsureEmpty:
                wb.Close(SaveChanges=False)
            else:
                if os.path.abspath(os.path.join(wb.Path, wb.Name)) == os.path.abspath(sPath):
                    if sTemplate=="":
                        return wb
                wb.Close(SaveChanges=True)
        
        if os.path.exists(sPath) and not bEnsureEmpty and sTemplate == "":
            Application.ScreenUpdating = False
            try:
                wb = Application.Workbooks.Open(sPath, UpdateLinks=bUpdateLinks)
                if wbActive is not None: wbActive.Activate()
            except BaseException as e:
                print(f"ERROR:Excel failed to open workbook:'{sPath}'")
                raise
            finally:
                Application.ScreenUpdating = bScreenUpdating
        else:
            bTemplateFileExists = os.path.exists(sTemplate)
            if bTemplateFileExists:
                bTemplateIsOpen = os.path.basename(sTemplate).lower() in set([wb.Name.lower() for wb in Application.Workbooks])
                if bTemplateIsOpen:
                    wbTemplate = Application.Workbooks[os.path.basename(sTemplate)]
                    if os.path.abspath(sTemplate) == os.path.abspath(wbTemplate.FullNameURLEncoded):
                        bReopenTemplate = True
                        wbTemplate.Close(SaveChanges=True)
            
          
            Application.ScreenUpdating = False
            try:
                wb = Application.Workbooks.Add(sTemplate if bTemplateFileExists else "")
                if wbActive is not None: wbActive.Activate()
            finally:
                Application.ScreenUpdating = bScreenUpdating
            
            sExt = sPath.split(".")[-1].strip().lower()
            FileFormat = {
                "xlsx"  : 51, #Application.xlWorkbookDefault,
                "xlsm"  : 52, #Application.xlOpenXMLWorkbookMacroEnabled,
                "xlsb"  : 50,
                "csv"   : 6 , #Application.xlCSV,
            }[sExt]
            
            cls.SafeGetWorkbook_SafeSaveAs(wb, sPath, FileFormat, excel=excel)
            
            if bReopenTemplate:
                Application.ScreenUpdating = False
                try:
                    Application.Workbooks.Open(sTemplate)
                    if wbActive is not None: wbActive.Activate()
                finally:
                    Application.ScreenUpdating = bScreenUpdating

        return wb
        
    @classmethod
    def SafeStripVBACode(cls, wb):
        VBComps = wb.VBProject.VBComponents
        vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document = 1, 2, 3, 100
        for VBComp in VBComps:
            if VBComp.Type in {vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm}:
                VBComps.Remove(VBComp)
            elif VBComp.Type == vbext_ct_Document:
                VBComp.CodeModule.DeleteLines(1, VBComp.CodeModule.CountOfLines)
            else:
                print("WARNING:Unexpected value for VBComp.Type:%r" % (VBComp.Type)) 
    
    @classmethod
    def ExistingVBComponent(cls, wb, name):
        ret = None
        try:
            ret = wb.VBProject.VBComponents(name)
        except BaseException as e:
            if str(e.args).find("Subscript out of range") < 0:
                raise  
        return ret
                
    @classmethod
    def SafeGetVBClassModulePublicNotCreatable(cls, wb, name, clear=True):
        ret = cls.ExistingVBComponent(wb, name)     
        cls_file_path = os.path.join(file__fileSysD, f"{str(uuid.uuid4())}.cls")                    
        expected_cls_file_code_header = f"""
VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
"""        
        if ret is None:
            cls_file_code = expected_cls_file_code_header
        else:
            ret.Export(cls_file_path)        
            wb.VBProject.VBComponents.Remove(ret)  
            with open(cls_file_path, "r") as ff:
                cls_file_code = ff.read()
                
            # Make sure our header expected_cls_file_code_header is still relevant
            expected_header_line_set = set([s.strip() for s in expected_cls_file_code_header.strip().split('\n') if s.strip() != ''])
            cls_file_code_split = cls_file_code.split('\n')
            found = False
            for k, s in enumerate(cls_file_code_split):
                s = s.strip()
                if s == 'Attribute VB_Exposed = False':
                    found = True
                    cls_file_code_split[k] = 'Attribute VB_Exposed = True'
                    expected_header_line_set.remove('Attribute VB_Exposed = True')
                if s in expected_header_line_set:
                    expected_header_line_set.remove(s)
                    if len(expected_header_line_set) == 0:
                        break
            if len(expected_header_line_set) > 0:
                print(f"WARNING:SafeGetVBClassModulePublicNotCreatable:len(expected_header_line_set) > 0:{sorted(expected_header_line_set)}")
                
            if found:
                cls_file_code = '\n'.join(cls_file_code_split)
                
        try:                
            with open(cls_file_path, "w") as ff:
                ff.write(cls_file_code)
            ret = wb.VBProject.VBComponents.Import(cls_file_path)            
        finally:
            if os.path.isfile(cls_file_path):
                os.remove(cls_file_path)  
        if clear:
            ret.CodeModule.DeleteLines(1, ret.CodeModule.CountOfLines)                
  
        return ret                

    @classmethod
    def SafeGetVBComponent(cls, wb, name, clear=True, vbext_ComponentType=None):
        if vbext_ComponentType is None:
            vbext_ComponentType = cls.vbext_ct_MSForm
        elif vbext_ComponentType not in cls.vbext_ct_ComponentTypes:
            raise Exception("vbext_ComponentType not in cls.vbext_ct_ComponentTypes")
        ret = cls.ExistingVBComponent(wb, name)
        if ret is None:
            ret = wb.VBProject.VBComponents.Add(vbext_ComponentType)
            try:
                ret.Name = name
            except BaseException as e:
                raise Exception(f"Failure running ret.Name = name:{vbext_ComponentType}:{name}:{e.__class__}:{e}")    
        else:
            if clear:
                if hasattr(ret, "Designer") and ret.Designer is not None:
                    ret.Designer.Controls.Clear()
                ret.CodeModule.DeleteLines(1, ret.CodeModule.CountOfLines)
        return ret                     
        
    @classmethod
    def SafeDelVBComponent(cls, wb, name):
        ret = cls.ExistingVBComponent(wb, name)
        if ret is not None:
            wb.VBProject.VBComponents.Remove(ret)
            
    @classmethod        
    def VBACodeEqual(cls, s1, s2):
        s1 = [s.strip().lower() for s in s1.split('\n') if s.strip() != '']
        s2 = [s.strip().lower() for s in s2.split('\n') if s.strip() != '']
        if len(s1) != len(s2):
            ret = False
        else:
            ret = True
            for ss1, ss2 in zip(s1, s2):
                if ss1 != ss2:
                    ret = False
                    break
        return ret
    
    @classmethod
    def SafeAddVBACompleteCodeModule(cls, wb, mod_name, code, vbext_ComponentType=None):
        if vbext_ComponentType is None:
            vbext_ComponentType = cls.vbext_ct_StdModule
        elif vbext_ComponentType not in cls.vbext_ct_ComponentTypes:
            raise Exception("vbext_ComponentType not in cls.vbext_ct_ComponentTypes")              
        vbcx = cls.ExistingVBComponent(wb, mod_name)
        if vbcx is None or vbcx.CodeModule.CountOfLines == 0 or not cls.VBACodeEqual(vbcx.CodeModule.Lines(1, vbcx.CodeModule.CountOfLines), code):
            if cls.vbext_ct_ClassModule == vbext_ComponentType:
                vbcx = cls.SafeGetVBClassModulePublicNotCreatable(wb, mod_name)
            else:
                vbcx = cls.SafeGetVBComponent(wb, mod_name, vbext_ComponentType = vbext_ComponentType)                
            vbcx.CodeModule.InsertLines(1, code)  
        return vbcx
            
    @classmethod        
    def SafeEnsurePythonVBAExtras(cls, wb):
        ret = "clsPythonVBAExtras"
        code = f"""Option Explicit
        
Function DebugPrint(msg)
Debug.Print msg
End Function

Function DoEventss()
DoEvents
End Function

Function MsgBoxx(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = vbNullString)
MsgBoxx = MsgBox(prompt, buttons, title)
End Function

Function UserFormShows(frm, Optional ByVal modal As Boolean = False)
call frm.Show(modal)
End Function

Function UserFormHides(frm)
call frm.Hide()
End Function

Function ThisWorkbooks()
Set ThisWorkbooks = ThisWorkbook
End Function

Function Selections()
Set Selections = Selection
End Function

"""
        vbcx = cls.ExistingVBComponent(wb, ret)    
        if vbcx is None or not cls.VBACodeEqual(vbcx.CodeModule.Lines(1, vbcx.CodeModule.CountOfLines), code):
            vbcx = cls.SafeGetVBClassModulePublicNotCreatable(wb, ret)
            vbcx.CodeModule.InsertLines(1, code)   
        return ret
            
    @classmethod        
    def SafeAttachVBAGlobalComObj(cls, wb, com_obj, mod_name=None):
        # Set the COM object 'com_obj' to be a public variable of a module named 'mod_name' in the VBA Project of wb 
        
        if mod_name is None: mod_name = f"mod{com_obj._username_.split('.')[1]}" # com_obj._username_ should be something like 'PythonComX.PythonComMSForms' 
        
        pclc = cls.SafeEnsurePythonVBAExtras(wb)
        
        code = f"""Public com_obj

Function Setter(com_obj_)
' To be called from Python to "set" the com_obj value
Set com_obj =  com_obj_
End Function

Function Getter()
' Intended to be called from Python to "get" the com_obj value ... the VBA code can simply reference "{mod_name}.com_obj"
Set Getter = com_obj
End Function

Function CreatePythonVBAExtras()
Set CreatePythonVBAExtras = New {pclc}
End Function

        """    
        vbcx = cls.SafeAddVBACompleteCodeModule(wb, mod_name, code)   
        wb.Parent.Application.Run(f"{mod_name}.Setter", com_obj)     
        return mod_name
            
                
    @classmethod
    def AddBooleanValidation(cls, rng, max_check_boxes_to_add=20):
        ws = rng.Parent    
        wb = ws.Parent
        excel = wb.Parent
        ScreenUpdating = excel.ScreenUpdating
        try:
            excel.ScreenUpdating = False
            if max_check_boxes_to_add > 0:
                cb = ws.CheckBoxes()
                for k, c in enumerate(rng.Cells):
                    if k >= max_check_boxes_to_add:
                        break
                    try:
                        x = cb.Add(c.Left, c.Top, c.Width, c.Height)
                        x.Caption = ""
                        x.LinkedCell = str(c.AddressLocal).replace("$", "")
                        #x.Value = xlOff '
                    except BaseException as e:
                        print(e)
                        
            vld = rng.Validation
            vld.Delete()
            vld.Add(Type=ExcelVBA.constants.xlValidateList, AlertStyle=ExcelVBA.constants.xlValidAlertStop, Operator=ExcelVBA.constants.xlBetween, Formula1="FALSE, TRUE")
            vld.IgnoreBlank = True
            vld.InCellDropdown = True
            vld.InputTitle = ""
            vld.ErrorTitle = ""
            vld.InputMessage = ""
            vld.ErrorMessage = ""
            vld.ShowInput = True
            vld.ShowError = True
        finally:
            excel.ScreenUpdating = ScreenUpdating
            
            
    class WorkbookOpenWithDisabledMacros:

        msoAutomationSecurityLow = 1 #(Enables all macros—not recommended for security reasons.)        
        msoAutomationSecurityByUI = 2 #(Uses the security setting chosen in the UI.)        
        msoAutomationSecurityForceDisable = 3 #(Disables all macros when opening a file via automation.)
        
        def __init__(self, wb):
            self.wb, self.full_path, self.xl, self.AutomationSecurity = wb, os.path.join(wb.Path, wb.Name), wb.Parent, None
        
        def __enter__(self):
            print(f"WorkbookOpenWithDisabledMacros:Closing and saving:{self.full_path} ...")
            self.wb.Close(SaveChanges:=True)
            self.wb = None
            self.AutomationSecurity_bak = self.xl.AutomationSecurity
            try:
                print(f"WorkbookOpenWithDisabledMacros:Disable VBA & Open:{self.full_path} ...")
                self.xl.AutomationSecurity = self.msoAutomationSecurityForceDisable
                self.wb = self.xl.Workbooks.Open(self.full_path)            
            except BaseException as e:
                print(f"ERROR:WorkbookOpenWithDisabledMacros:Failure reopening Workbook:{e.__class__}:{e}:{self.full_path}")
                self.wb = None
                try:
                    self.xl.AutomationSecurity = self.AutomationSecurity_bak
                except BaseException as ee:
                    print(f"WARNING:WorkbookOpenWithDisabledMacros:Failed to restore Excel.AutomationSecurity:{ee.__class__}:{ee}")
                    self.AutomationSecurity_bak = None
                raise e
            return self 

        def __exit__(self, exc_type, exc_value, traceback):
            if self.AutomationSecurity_bak is not None:
                try:
                    self.xl.AutomationSecurity = self.AutomationSecurity_bak
                except BaseException as ee:
                    print(f"WARNING:WorkbookOpenWithDisabledMacros:Failed to restore Excel.AutomationSecurity:{ee.__class__}:{ee}")
            if self.wb is not None:
                print(f"WorkbookOpenWithDisabledMacros:Closing and saving:{self.full_path} ...")
                self.wb.Close(SaveChanges:=True)
                self.wb = None
            try:
                print(f"WorkbookOpenWithDisabledMacros:Reloading Workbook:{self.full_path} ...")
                self.wb = self.xl.Workbooks.Open(self.full_path)
                print(f"WorkbookOpenWithDisabledMacros:Workbook reload OK:{self.full_path} Complete")
            except BaseException as ee:
                print(f"WARNING:WorkbookOpenWithDisabledMacros:Failure reloading Workbook:{ee.__class__}:{ee}:{self.full_path}")
            
            return False  
            
    vba_header_end_pattern = re.compile(r"^(Attribute\s+VB_[^\n]+\n)+(?!Attribute)", re.MULTILINE)            
            
    @classmethod
    def remove_vba_header(cls, vba_text):
        match = cls.vba_header_end_pattern.search(vba_text)
        return vba_text[match.end():].lstrip() if match else vba_text
            

            
                
class VBAIO:            
   
    
    _vbaio_lck = threading.Lock()
    _folder_base_name_2_vbext_ct_ComponentTypesx = None
    
    @classproperty
    def folder_base_name_2_vbext_ct_ComponentTypesx(cls):
        if cls._folder_base_name_2_vbext_ct_ComponentTypesx is None:
            with cls._vbaio_lck:
                if cls._folder_base_name_2_vbext_ct_ComponentTypesx is None:
                    cls._folder_base_name_2_vbext_ct_ComponentTypesx = {v[0] : k for k, v in ExcelVBA.vbext_ct_ComponentTypesx.items()}
        return cls._folder_base_name_2_vbext_ct_ComponentTypesx
        
    @classproperty
    def vba_root(cls):
        return os.path.join(file__parentDr, "src", "vba")
        
    @classmethod    
    def VBAIterateFiles(cls):
        for r, d, f in os.walk(cls.vba_root):
            bn = os.path.basename(r) 
            if bn in cls.folder_base_name_2_vbext_ct_ComponentTypesx:
                Type = cls.folder_base_name_2_vbext_ct_ComponentTypesx[bn]
                extn = f".{ExcelVBA.vbext_ct_ComponentTypesx[Type][1]}"
                for p in f:
                    if p.endswith(extn):
                        Name = p[:-len(extn)]
                        yield Type, Name, os.path.join(r, p)
                        
    @classmethod    
    def VBAIterateComps(cls, wb):                        
        VBComps = wb.VBProject.VBComponents
        for VBComp in VBComps:
            if VBComp.Type not in ExcelVBA.vbext_ct_ComponentTypesx:
                raise Exception("VBComp.Type not in ExcelVBA.vbext_ct_ComponentTypesx")
            yield VBComp
            
    @classmethod
    def VBAExport(cls, wb):
        code_name_2_comp_name = {ws.CodeName: ws.Name for ws in wb.Sheets}
        code_name_2_comp_name["ThisWorkbook"] = "ThisWorkbook"
        for VBComp in cls.VBAIterateComps(wb):
            if VBComp.Type != ExcelVBA.vbext_ct_Document:
                base_name = VBComp.Name
            else:
                if VBComp.CodeModule.CountOfLines == 0:
                    continue
                base_name = code_name_2_comp_name[VBComp.Name]
            file_path = os.path.join(cls.vba_root, ExcelVBA.vbext_ct_ComponentTypesx[VBComp.Type][0], f"{base_name}.{ExcelVBA.vbext_ct_ComponentTypesx[VBComp.Type][1]}")
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            print(f"VBAExport:{file_path}")
            VBComp.Export(file_path)    
            
    @classmethod
    def VBAImport(cls, wb):
        wowdm = ExcelVBA.WorkbookOpenWithDisabledMacros(wb)
        with wowdm:
            wb = wowdm.wb
            for Type, Name, fp in cls.VBAIterateFiles():          
                print(f"VBAImport:%3d %-30s %s" % (Type, Name, fp))                       
                if Type != ExcelVBA.vbext_ct_Document:
                    # The most general way to do this is to delete the component and reload it
                    VBComp = ExcelVBA.ExistingVBComponent(wb, Name)                          
                    if VBComp is not None:
                        wb.VBProject.VBComponents.Remove(VBComp)
                    wb.VBProject.VBComponents.Import(fp)
                else:
                    # This code is associated with a sheet ... we can't use VBComponents.Import
                    if Name == "ThisWorkbook":
                        VBComp = ExcelVBA.ExistingVBComponent(wb, "ThisWorkbook")
                    else:
                        ws = ExcelVBA.SafeGetWorksheet(wb, Name)
                        VBComp = ExcelVBA.ExistingVBComponent(wb, ws.CodeName)
                        
                    if VBComp.CodeModule.CountOfLines > 0 :
                        VBComp.CodeModule.DeleteLines(1, VBComp.CodeModule.CountOfLines)
                    with open(fp, "r") as ff:
                        code = ExcelVBA.remove_vba_header(ff.read())
                    VBComp.CodeModule.AddFromString(code)
        return wowdm.wb
   
   
if True and __name__ == "__main__":
    
    xl = ExcelVBA.get_excel_x()
    wb = xl.Workbooks("rosehaven_florist_junk.xlsm")
    #wb = xl.Workbooks("Journal.xlsm")
    #VBAExport(wb)
    wb = VBAIO.VBAImport(wb)
    VBAIO.VBAExport(wb)
    


    
