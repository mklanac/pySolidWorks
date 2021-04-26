# -*- coding: utf-8 -*-
"""
Created on Thu Oct  3 09:50:54 2019

@author: Mario
"""

import os
import win32com.client
import pythoncom
import numpy as np

class Solidworks():
    def __init__(self):
        """
        Initialize connection between Solidworks and Python
           Parameters:
               none
           Returns:
               SW (class)
        """
        
        self.swYearLastDigit = 9  # Solidworks 2016
        self.swApp = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(self.swYearLastDigit-2)))  # e.g. 20 is Solidworks 2012,  24 is Solidworks 2016
        self.swApp.Visible = 1
        
        
    def GenerateNewModel(self, motor):
        
        eqnMgr = self.model.GetEquationMgr
        NoneSWvariableParameters = len(motor.variableParameters)    # J, Dout, l, hs, bt, ...
        for i, var in enumerate(motor.SolidworksDataSorted[0, :]):
            #print(var)
            if '@Sketch' in var:
                myDimension = self.model.Parameter(var)
                varScalingFactor = motor.SolidworksDataSorted[2, i]
                if isinstance(varScalingFactor, str):
                    if var=='Dr@Sketch1':
                        myDimension.SystemValue = motor.geometryData['Dr']/1000
                    else:
                        myDimension.SystemValue = motor.Parameters[varScalingFactor]*motor.motorInput['X'][i+NoneSWvariableParameters]
                else:
                    myDimension.SystemValue = varScalingFactor*motor.motorInput['X'][i+NoneSWvariableParameters]
                #print(myDimension.SystemValue)
            else:
                eqnMgr.Equation(motor.SolidworksDataSorted[1, i], var+'='+str(motor.motorInput['X'][i+NoneSWvariableParameters]))   # index number of equation in Solidworks and variable name with value as a second argument, e.g., # eqnMgr.Equation(0, 'dbp1=0.1')
                #print(str(motor.motorInput['X'][i+NoneSWvariableParameters]))
            self.model.EditRebuild3
                
                
                
    def GetValueOfDimension(self, DimensionName):
        """
        Get sketch dimensions
           Parameters:
               Soldworks(self) (class)
               DimensionName (str) : name of dimension, e.g., DimensionName='Dr@Sketch1'
           Returns:
               dimension value (default: angles are returned in radians and length in meters)
        """
        myDimension = self.model.Parameter(DimensionName)
        
        return myDimension.SystemValue
    
    
    def GetStudy(self):
        """
        Allows access to the first study.
           Parameters:
               Soldworks(self) (class)
           Returns:
               Study
        """
        CWObject = self.swApp.GetAddInObject("SldWorks.Simulation")
        COSMOSWORKS = CWObject.COSMOSWORKS
        ActDoc = COSMOSWORKS.ActiveDoc
        StudyMngr = ActDoc.StudyManager
        Study = StudyMngr.GetStudy(0)
        
        return Study
    
    
    def SetForce(self, Study, forceName, ForceValue):
        """
        Set forces on an object
           Parameters:
               Soldworks(self) (class)
               Study (class 'win32com.client.CDispatch')- allows access to a study
               forceName (string(one force) or list(more forces on object)) : name of force applied on an object, i.e., 'F1', 'F2', ...
               forceValue (float or list) : value/s of force
           Returns:
               ret (bool) : True(succesfully applied all forces) or False (somethinf failed)
        """
        LBCMgr = Study.LoadsAndRestraintsManager
        errCode= win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)
        myForce = LBCMgr.GetLoadsAndRestraints(0, errCode)
        
        for force in range(LBCMgr.Count):
            try:
                myForce = LBCMgr.GetLoadsAndRestraints(force, errCode)
                if isinstance(forceName, str):
                    if forceName==myForce.Name:
                        myForce.ForceBeginEdit
                        myForce.NormalForceOrTorqueValue = ForceValue
                        myForce.ForceEndEdit
                        break
                elif isinstance(forceName, (np.ndarray, np.generic)):
                    forceName = forceName.tolist()
                elif isinstance(forceName, list):
                     if myForce.Name in forceName:
                         myForce.ForceBeginEdit()
                         myForce.NormalForceOrTorqueValue = ForceValue[forceName.index(myForce.Name)]
                         myForce.ForceEndEdit
            except AttributeError as error:
                print('There was a problem setting the force on an object.', )
        
    def ConfigureMeshAndRunAnalysis(self, Study):
        """
        Set mesh quality and run analysis
        Parameters:
            Soldworks(self) (class)
            Study (class 'win32com.client.CDispatch')- allows access to a study
        Returns:
            ret (bool) : True(succesfully applied all forces) or False (somethinf failed)
        """
        CwMesh = Study.Mesh
        CwMesh.Quality = 1        # 1 = high quality elements
        CwMesh.MesherType = 1     # 1 = curvature-based
        CwMesh.GrowthRatio = 1.6
        CwMesh.MinElementsInCircle = 8
        el= win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, -1)   # element size
        tl= win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.05) # tolerance
        CwMesh.GetDefaultElementSizeAndTolerance(0, el, tl)
        ret = Study.CreateMesh(0, el.value, 0.05)
        ret = Study.RunAnalysis
        
        return ret
    
    def GetMinMaxStress(self, Study):
        """
        Gets the algebraic minimum and maximum for the von Mises stress component (first argument -> 9), 0 element (second argument) and solution step (third argument -> Static ())
        Parameters:
            Soldworks(self) (class)
            Study (class 'win32com.client.CDispatch')- allows access to a study
        Returns:
            Stress (tuple) : tuple in the following format (node_with_minimum_stress, minimum_stress_value, node_with_maximum_stress, maximum_stress_value)
        """
        results = Study.Results
        errCode= win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)
        arg1= win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        
        Stress = results.GetMinMaxStress(9, 0, 1, arg1, 0 , errCode)
        
        return Stress


    def RunMacro(self, MacroPath):
        """
        Run macro
           Parameters:
               Soldworks(self) (class)
               MacroPath (str) : full path of the macro
           Returns:
               none
               DXF file is exported to DXF folder
        """
        arg= win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)
        self.swApp.RunMacro2(MacroPath,"Module", "main",1, arg)
        
    def NewDoc(self, modelPath, docType):
        """
        Create new document at the location modelPath
           Parameters:
               Soldworks(self) (class)
               modelPath (str) : full path of the model
               docType (str) : 'Part' or 'Assembly' 
           Returns:
               none
        """
        self.modelPath = modelPath
        template = 'C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2016\\templates\%s.drwdot' % docType
        self.model = self.swApp.NewDocument(template, 0, 0, 0)
        
        self.SaveDoc(modelPath)


    def SaveDoc(self, modelPath):
        """
        Save document at the location modelPath
           Parameters:
               Soldworks(self) (class)
               modelPath (str) : full path of the model
           Returns:
               none
        """
        self.modelPath = modelPath
        self.model = self.swApp.ActiveDoc.SaveAs3(modelPath, 0, 2)
        
    
    def OpenDoc(self, modelPath):
        """
        Open document at the location modelPath
           Parameters:
               Soldworks(self) (class)
               modelPath (str) : full path of the model
           Returns:
               none
        """
        self.modelPath = modelPath
        self.model = self.swApp.OpenDoc(modelPath, 1)
        
    def CloseDoc(self):
        """
        Close document without saving
           Parameters:
               Soldworks(self) (class)
           Returns:
               none
        """
        self.swApp.CloseDoc(self.modelPath)