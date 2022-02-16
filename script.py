# -*- coding: utf-8 -*-
"""
Created on Sat Sep 26 04:50:18 2020

@author: sluke
"""

import pandas as pd
import numpy as np
import win32com.client # this program only works on Windows. I used Bootcamp on my Mac to create a Windows partition.

# bringing in database
df = pd.read_csv(r"C:\Users\sluke\Desktop\fake_recruit_db.csv")
# making adjustment to data frame
df["pref_name"] = df["pref_name"].fillna(0) # Filling all "pref_names" to 0 if they dont have a preferred name. 

print(df.head())
print(df.columns)

# trial database
tdf = df.loc[0:1,:] # trial to run through while testing
#print(tdf.loc[:,"first_name":"last_name"])

# seach database
sdf = df.loc[df.first_name == "Clay",:]
#print(sdf.loc[:,"first_name":"last_name"])

#%%
# bringing in photoshop file
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"C:\Users\sluke\Desktop\JerseyPile.psd")
doc = psApp.Application.ActiveDocument

#%%
##################
###    TEXT    ###
##################
# Change layers with text such as names and numbers. You can also adjust size of font, warp style, etc based on length of name.

# Text Options:
        # size = 0 - 250 (character size)
        # FauxBold = True/False | FauxItalics = True/False | Underline  = [1,2,3] (style)
        # Direction = [1 (horizontal text), 2 (vertical text)]
        # HorizontalScale = 0-1000 (character scaling))
        # Tracking = -1000 - 10000 (space between letters))
        # VerticalScale = 0 - 1000 poroportion to horizontal scale)
        # WarpStyle = [1 (normal), 2 (arch), etc]
        # WarpBend = -100 - 100 amound of bend
        
def firstNameChange(layer_name):
    layerText = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in photoshop folder
    text_of_layer = layerText.TextItem
    if pref_name == "0": # it no "pref_name" will revert to first_name
    # pref_name adjustments to "0" made adjustments to dataframe
        text_of_layer.contents = f'{first_name}' 
    else:
        text_of_layer.contents = f'{pref_name}' # new text entered in layer
    return text_of_layer.contents

def lastNameChange(layer_name):
    layerText = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in photoshop folder
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f'{last_name}' # new text entered in layer
    # Changing Text Size (Need to list other text adjustments that can be made)
    if len(text_of_layer.contents) <= 14:
        text_of_layer.size = 16.45
    else:
        text_of_layer.size =10
    return text_of_layer.contents

def fullNameChange(layer_name):
    layerText = doc.layerSets["Template 1"].ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in folder
    text_of_layer = layerText.TextItem
    ### Changing Text Contents
    if pref_name == "0": # it no "pref_name" will revert to first_name
    # pref_name adjustments to "0" made adjustments to dataframe
        text_of_layer.contents = f'{first_name} {last_name}' 
        #if len(text_of_layer.contents) <= 10:
         #   text_of_layer.size = 16.45
        #else:
        #    text_of_layer.size =10
    else:
        text_of_layer.contents = f'{pref_name} {last_name}' # new text entered in layer
    
# =============================================================================
#     if len(text_of_layer.contents) <= 10:
#         text_of_layer.size = 16   
#     elif len(text_of_layer.contents) <= 14:
#         text_of_layer.size =14
#     else:
#         text_of_layer.size = 12
# =============================================================================

    return text_of_layer.contents

def numberChange(layer_name):
    layerText = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in folder
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f'{number}' # new text entered in layer
    return text_of_layer.contents

def jerseyNameChange(layer_name):
    layerText = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in folder
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f'{last_name}' # new text entered in layer
    # Changing Text Size 
	if len(text_of_layer.contents) <= 10: # if the last name is ten or less characters the size of font is 15
		text_of_layer.size = 15
	else:
		text_of_layer.size = 10 # if the last name is over ten characters the size of the font is ten
    return text_of_layer.contents

def jerseyNumberChange(layer_name):
    layerText = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in folder
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f'{number}' # new text entered in layer
    return text_of_layer.contents


##################
####  LAYERS  ####
##################

def layerControl(layer_name, variable_name): # havent tested this function yet
    layerPs = doc.layerSets[layer_name] # name of layer in PS might need to add layerSets if in folder
    if variable_name == "fair" or variable_name == "olive":
        variable_name = "W"
    elif variable_name == "light_brown":
        variable_name = "M" 
    elif variable_name == "brown":
        variable_name = "MD"  
    elif variable_name == "black_brown" or variable_name == "dark_brown":
        variable_name = "D"        
    if layer_name == variable_name: # "t" = True, but can do binary system as well
        layerPs.visible = True
    else:
        layerPs.visible = False
 
def skinToneControl(layer_name, variable_name): # havent tested this function yet
    layerPs = doc.layerSets[layer_name] # name of layer in PS might need to add layerSets if in folder
    if variable_name == "fair" or variable_name == "olive":
        variable_name = "W"
    elif variable_name == "light_brown":
        variable_name = "M" 
    elif variable_name == "brown":
        variable_name = "MD"  
    elif variable_name == "black_brown" or variable_name == "dark_brown":
        variable_name = "D"        
    if layer_name == variable_name: # "t" = True, but can do binary system as well
        layerPs.visible = True
    else:
        layerPs.visible = False
        
def bodyTypeControl(layer_name, variable_name): 
    layerPs = doc.layerSets[layer_name] # name of layer in PS might need to add layerSets if in folder     
    if layer_name == variable_name: # "t" = True, but can do binary system as well
        layerPs.visible = True
    else:
        layerPs.visible = False
        
# =============================================================================
#         
# def hideLayer(layer_name):
#     layerPs = doc.ArtLayers[f"{first_name}_{last_name}"] # name of layer
#     if not_visible == "yes": # "t" = True, but can do binary system as well
#         layerPs.visible = False
#     else:
#         layerPs.visible = True
# =============================================================================
        
def hideLayer(layer_name): # Used to delete unecessary layers
    layerPs = doc.layers[layer_name]  # name of layer in PS might need to add layerSets if in folder
    layerPs.visible = False  
        
def deleteLayer(layer_name): # Used to delete unecessary layers
    layerPs = doc.layers[layer_name]  # name of layer in PS might need to add layerSets if in folder
    layerPs.Delete()
    
  
#%%

##################
####  PHOTOS  ####
##################

# Layer Options:
    #photoLayer.AdjustBrightnessContrast(100,100)
    #photoLayer.ApplyGaussianBlur(5)
    #photoLayer.Desaturate() #grayscales layer
    #photoLayer.Posterize(200  ) 
    #photoLayer.BlendMode = 1-30 (normal = 2) [pg. 159 in manual]
    #photoLayer.ApplyDisplace()
    #photoLayer.Resize(120,120) #resizes image by percentage doubles width
    #photoLayer.ShadowHighlight(80,20,20,20,20,20,20,20,20,20) #supposed to add shadow
    #photoLayer.Threshold(175)
    #photoLayer.AppyOffset(500,-500)
    #photoLayer.ApplyShear((50,-50),(-50,50))
    #photLayer.Mixchannesl()
    # LAYER SET ADJUSTMENTS
    #layerSet.Visible() # you can hide entire layerSets

def insertHeadSwap(layer_name, delta_x = 0, delta_y = 0): #
    # make sure that the layer you bring it in as is one larger than
    # the highest empty Layer.
    # also put highlight layer that is hidden before running the program
    # enter pathway to destination of photo
    psApp.Open(f"C:/Users/sluke/Desktop/head_swaps/{first_name}_{last_name}.png")
        # switch acive doucment to doc2
    doc2 = psApp.Application.ActiveDocument
    doc2.ArtLayers("Layer 1")
        # finding pixels of heights and width for head swaps hopefully they are done 
        # on different ratio files
    doc2Height = int(doc2.Height)
    doc2Width = int(doc2.Width)      
    print(f"{first_name} {last_name}\n")
    print(f"Height: {doc2Height}  \nWidth: {doc2Width} \nRatio: {doc2Width/doc2Height}\n")
        # based on the raio can change size of photo to get similar sizing
    if doc2Width/doc2Height <= .37: # skill holding ball
        doc2.ResizeImage(doc2Width*.61)   
    elif doc2Width/doc2Height < .44: # big skill arms cross
        doc2.ResizeImage(doc2Width*.42)
    elif doc2Width/doc2Height <= .55: # skill hold collar
        doc2.ResizeImage(doc2Width*.40)      
    elif doc2Width/doc2Height <= .6: # lineman arms cross
        doc2.ResizeImage(doc2Width*.60)    
    elif doc2Width/doc2Height <= .65: # lineman arms down
        doc2.ResizeImage(doc2.Width*.50)         
    else: # big skill holding hands 
        doc2.ResizeImage(doc2Width*.47)
        
    doc2.Trim() # Trims blank space around the cutout allows us to get
        # closer measurement so we can standardized height for more accureate
        # picture placement
    trimHeight = doc2.Height
    trimWidth = doc2.Width
    print(f"Height: {trimHeight}  \nWidth: {doc2Width} \nRatio: {trimWidth/doc2Height}\n")
    
        # creates dictionary for each individual document height
    if doc2Height not in d:
        d[doc2Height] = int(trimHeight)
        dlist.append(doc2Height)
        # selecting everything in doc2 and copying
    doc2.Selection.SelectAll()
    doc2.Selection.Copy()

        
    doc2.Close(2) # closes doc2 with out saving changes (2)
        # reselecting original layer
    doc = psApp.Application.ActiveDocument
    doc.ActiveLayer = doc.Layers["Active_Select"] # important to select
        # wanted layer everytime allows naming to stay consistent
        # put object in layer so it doesnt past picture there
        # then change opacity and fill to 0%
    doc.Paste()
        

    blankLayer = doc.ArtLayers[layer_name] # name of layer in PS might need to add layerSets if in folder
        #blankLayer = doc.layerSets["Check"].ArtLayers[layer_name] # .layerSets["Check"] 
        # .layerSets["Check"] allows you to go into folders need to include in
        # hiding photo below. Messes up when adding next photo.
        
    photoLayer = doc.ArtLayers["Layer 1"] # higest Layer is 9 in doc
    photoLayer.Name = (f'{first_name}_{last_name}') # naming photo layer this way so we can hide it later
    photoLayer.MoveBefore(blankLayer) # make sure path in blankLayer (above)
        # is correct. other options (MoveAfter, MoveFirst?, MoveLast?)

        # Add any adjusments
    
    
        # translate layer left(negative number) and up(negative number)       
    if doc2Width/doc2Height <= .37: # skill holding ball
        photoLayer.Resize(100*((d[4528]-trimHeight+d[4528])/d[4528]),100*((d[4528]-trimHeight+d[4528])/d[4528]))
        photoLayer.Translate(300,500)   
    elif doc2Width/doc2Height < .42: # big skill arms cross
        photoLayer.Resize(100*((d[4758]-trimHeight+d[4758])/d[4758]),100*((d[4758]-trimHeight+d[4758])/d[4758]))
        photoLayer.Translate(300,500)
    elif doc2Width/doc2Height <= .55: # skill holding collar
        photoLayer.Resize(100*((d[4593]-trimHeight+d[4593])/d[4593]),100*((d[4593]-trimHeight+d[4593])/d[4593]))
        photoLayer.Translate(300,500) 
    elif doc2Width/doc2Height <= .6: # lineman arms cross
        photoLayer.Resize(100*((d[3045]-trimHeight+d[3045])/d[3045]),100*((d[3045]-trimHeight+d[3045])/d[3045]))
        photoLayer.Translate(300,500)   
    elif doc2Width/doc2Height <= .65: # linemna arms down
        photoLayer.Resize(100*((d[3398]-trimHeight+d[3398])/d[3398]),100*((d[3398]-trimHeight+d[3398])/d[3398]))
        photoLayer.Translate(300,500)        
    else: # big skill holding hands
        photoLayer.Resize(100*((d[5472]-trimHeight+d[5472])/d[5472]),100*((d[5472]-trimHeight+d[5472])/d[5472]))
        photoLayer.Translate(300,500) # 


def dropShadow(layer_name): # Not Working
    layerPs = doc.layers[layer_name] # name of layer in PS might need to add layerSets if in folder
    layerPs.Duplicate()
    layerPs.AdjustBrightnessContrast(-100,-100)
    layerPs.AdjustBrightnessContrast(-100,-100)
    layerPs.ApplyGaussianBlur(15)
    layerPs.Opacity = 50
    layerPs.Translate(25,-25)
    layerPs.Name = f'{first_name}_{last_name}_shadow' # naming photo layer this way so we can hide it later
    layerPs2 = doc.layers[f"{first_name}_{last_name} copy"]
    layerPs2.Name = f"{first_name}_{last_name}"

#%%

##################
####   SAVE   ####
##################        
        
def saveFile(file_type): # enter png, jpg, or psd
# can order final products by creating subfolders. I found separating them by postion was helpful when sending to coaches
    if file_type == "png":   
        options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
        options.Format = 13
        options.PNG8 = False
        pngfile = f"C:/Users/sluke/Desktop/photo_dump/{pos}/{first_name}_{last_name}_{year}.png" # specify the path were you want the final product saved
        doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)
    elif file_type == "jpg": 
        options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
        options.Format = 6
        options.Quality = 100
        jpgFile = f"C:/Users/sluke/Desktop/photo_dump/{pos}/{first_name}_{last_name}_{year}.jpg" # specify the path were you want the final product saved
        doc.Export(ExportIn=jpgFile, ExportAs=2, Options=options)
    else: # psd file saving isnt working
        options = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
        psdfile = f"C:/Users/sluke/Desktop/{first_name}_{last_name}.psd"
        doc.SaveAs(SaveIn = psdfile, Options = options)
        
#%%
# provides list when photographs are missing
missing_list = []
# dictionary to hold photo dimensions
d = {}


for index, row in df.iterrows():
    # Variable Names
    first_name = row["first_name"]
    pref_name = row["pref_name"]
    last_name = row["last_name"]
    year = row["class"]
    pos = row["position"]
    number = row['number'] 
    
#%%
    # Cleaning data
    first_name = first_name.strip() # stripping any uncessary spaces    
    pref_name = str(pref_name) # changing from number to string to be able to strip uncessary spaces
    pref_name = pref_name.strip() # stripping any uncessary spaces    
    last_name = last_name.strip() # stripping any uncessary spaces 
    pos = pos.strip() # stripping any uncessary spaces
    number = str(number) # changing from number to string to be able to strip uncessary spaces
    number = number.strip() # stripping any uncessary spaces
    
#%%
    
    # Script
    fullNameChange("full name") # these have to match the name of layer in photoshop
    jerseyNameChange("NAME8") # these have to match the name of layer in photoshop
    jerseyNumberChange("NUMBER8") # these have to match the name of layer in photoshop
    saveFile("jpg")
  
     
print(d)    
print(dlist)


# try, except, and finally
# when dealing with inserting photos allows you to move on
# if file is missing, can also alert yourself to this occasion
# example below:
# =============================================================================
#     # Script
#     nameChange("name")
#     signature("signature")
#     changeNumber("number", 1, 1)
#     try:
#         insertPhoto("blanklayer",0,0) 
#         hideLayer("Active_Select") 
#         # hide selected active layer 
#         saveFile("jpg")
#         hidePhoto("yes")
#     #deleteLayer(f"{first_name}_{last_name}")
#     except:
#         missing_list.append(f"{first_name} {last_name}")
#         print("********************************************\n")
#         print(f"Can't attach photo for {first_name} {last_name}. Photo is either missing or not named properly.\n")
#         print("********************************************\n\n")
# 
#     finally:
#         changeNumber("number", -1, -1)
#       print("Missing photos for these people:")
#       print(missing_list)
#       print(d)
# =============================================================================
    
    

