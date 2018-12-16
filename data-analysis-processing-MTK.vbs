
Sub Form_OnLoad  
   'the path where the resulting xy files should be exported to   
exportPath = Analysis.Path  
   
   
'the list of masses for the extracted mass traces   
massesToExport= Array(586.218,602.213,618.210,616.192,632.187,474.165,147.047,440.178,425.150,567.200,731.250,601.210,439.170,457.180)   
   
Dim currentAnalysis   
   
For Each currentAnalysis in Application.Analyses   
 
    'The following example defines two selected ranges 
    '(10-15min and 20-22min) of the Chromatogram object 
    ' of the first chromatogram loaded and then calculates  
    'a profile spectrum placed in the Compound Mass List.  
    'Analysis.Chromatograms(1).AddRangeSelection 10, 15, 0, 0  
    'Analysis.Chromatograms(1).AddRangeSelection 20, 22, 0, 0  
    'Analysis.Chromatograms(1).AverageMassSpectrum false, true   
    Analysis.Chromatograms(1).AddRangeSelection 0.2, 0.22, 0, 0  
    Analysis.Chromatograms(1).AverageMassSpectrum false, true 
 
     
    'The following example defines two selected ranges 
    '(10-15min and 20-22min) of the Chromatogram object 
    'of the first chromatogram loaded: 
    'Analysis.Chromatograms(1).AddRangeSelection 0.2, 0.22, 0, 0  
 
 
    'The following example recalibrates the analysisinternally based on the 
    'recalibration of the first spectrum not being part of a compound:  
    'só funciona em modo positivo porque algures o método de MS para negativo não tem 
    'as definições completas do ESI formate neg 
 
    'implica que exista um compound spectra chamado "+MS, 0.2min" 
    Analysis.RecalibrateInternal   
    'nao da erro se nao houver 
  
 
       
    'Set BPC = CreateObject("DataAnalysis.BPCChromatogramDefinition")   
    'currentAnalysis.Chromatograms.AddChromatogram BPC  
      
    For Each mass in massesToExport   
          
        Set EIC = CreateObject("DataAnalysis.EICChromatogramDefinition")   
        EIC.range = mass   
        EIC.WidthLeft = 0.01  
        EIC.WidthRight = 0.01  
        currentAnalysis.Chromatograms.AddChromatogram EIC   
    Next   
   
    'Set TIC = CreateObject("DataAnalysis.TICChromatogramDefinition")   
    'currentAnalysis.Chromatograms.AddChromatogram TIC   
      
    'isto permite exportar cada cromatograma na analise para XY 
    'o que devia dar para fazer em excel, mas precisará de uma macro 
    'para anular Y em X-del e X+del  
    'For Each currentChromatogram in currentAnalysis.Chromatograms   
    '    currentChromatogram.Export exportPath+currentAnalysis.name+"_"+currentChromatogram.name+".xy", daXY   
    'Next   
    
    'Extract MS2 for all compounds in all chromatograms
    For each ONE in Analysis.Chromatograms
      ONE.FindCompounds
    Next
    

Next  
 
Analysis.Chromatograms(3).Name_ = "MTK+H,pos  586.218 witdh=0.01 +All MS" 
Analysis.Chromatograms(4).Name_ = "MTK+O+H,pos  602.213 witdh=0.01 +All MS" 
Analysis.Chromatograms(5).Name_ = "MTK+2O+H,pos  618.210 witdh=0.01 +All MS" 
Analysis.Chromatograms(6).Name_ = "MTK_COOH+H,pos  616.192 witdh=0.01 +All MS" 
Analysis.Chromatograms(7).Name_ = "MTK_COOH+O+H,pos  632.187 witdh=0.01 +All MS" 
Analysis.Chromatograms(8).Name_ = "MTK_Sdealk1+H,pos  474.165 witdh=0.01 +All MS" 
Analysis.Chromatograms(9).Name_ = "MTK_Sdealk2+H,pos  147.047 witdh=0.01 +All MS" 
Analysis.Chromatograms(10).Name_ = "MTK_Sdealk3+H,pos  440.178 witdh=0.01 +All MS" 
Analysis.Chromatograms(11).Name_ = "MTK IMP 425.150 witdh=0.01 +All MS" 
Analysis.Chromatograms(12).Name_ = "MTK IMP 567.200 witdh=0.01 +All MS" 
Analysis.Chromatograms(13).Name_ = "MTK IMP 731.250 witdh=0.01 +All MS" 
Analysis.Chromatograms(14).Name_ = "MTK IMP 601.210 witdh=0.01 +All MS" 
Analysis.Chromatograms(15).Name_ = "MTK IMP 439.170 witdh=0.01 +All MS" 
Analysis.Chromatograms(16).Name_ = "MTK IMP 457.180 witdh=0.01 +All MS" 
 
Analysis.Save
 
End Sub  
 
 
  
 
 
 
 
