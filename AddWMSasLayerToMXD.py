# snippets used from Mark Cederholm
import requests
import arcpy
import comtypes, os
import sys
import arcview
from comtypes import COMError


def main():
    pass


# Luke Pinner 2014 with alterations by Sean Eagles 2021
def CreateObject(COMClass, COMInterface):
    """ Creates a new comtypes POINTER object where
        COMClass is the class to be instantiated,
        COMInterface is the interface to be assigned
    """
    ptr = comtypes.client.CreateObject(COMClass, interface=COMInterface)
    return ptr


def CreateMXD(path):
    GetModule('esriCarto.olb')
    import comtypes.gen.esriCarto as esriCarto
    pMapDocument = CreateObject(esriCarto.MapDocument, esriCarto.IMapDocument)
    pMapDocument.New(path)
    pMapDocument.Save()  # probably not required...


def NewObj(MyClass, MyInterface):
    """Creates a new comtypes POINTER object where\n\
    MyClass is the class to be instantiated,\n\
    MyInterface is the interface to be assigned"""
    from comtypes.client import CreateObject
    try:
        ptr = CreateObject(MyClass, interface=MyInterface)
        return ptr
    except:
        return None


def CType(obj, interface):
    """Casts obj to interface and returns comtypes POINTER or None"""
    try:
        newobj = obj.QueryInterface(interface)
        return newobj
    except:
        return None


def GetLibPath():
    """Return location of ArcGIS type libraries as string"""
    # This will still work on 64-bit machines because Python runs in 32 bit mode
    import _winreg
    keyESRI = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\ESRI\\Desktop10.8")
    return _winreg.QueryValueEx(keyESRI, "InstallDir")[0] + "com\\"


def GetModule(sModuleName):
    """Import ArcGIS module"""
    from comtypes.client import GetModule
    sLibPath = GetLibPath()
    GetModule(sLibPath + sModuleName)


# Luke Pinner 2014 with alterations and extension by Sean Eagles 2021
def ConnectWMS(url):

    GetModule("esriSystem.olb")
    GetModule("esriCarto.olb")
    GetModule("esriGISClient.olb")

    import comtypes.gen.esriSystem as esriSystem
    import comtypes.gen.esriCarto as esriCarto
    import comtypes.gen.esriGISClient as esriGISClient

    # Create an WMSMapLayer Instance - this will be added to the map later
    pWMSGroupLayer = NewObj(esriCarto.WMSMapLayer, esriCarto.IWMSGroupLayer)

    # Create and configure wms connection name, this is used to store the
    # connection properties
    pPropSet = NewObj(esriSystem.PropertySet, esriSystem.IPropertySet)
    pPropSet.SetProperty("URL", url)
    pConnName = NewObj(
        esriGISClient.WMSConnectionName,
        esriGISClient.IWMSConnectionName)

    pConnName.ConnectionProperties = pPropSet

    # Use the name information to connect to the service
    pDataLayer = CType(pWMSGroupLayer, esriCarto.IDataLayer)
    pName = CType(pConnName, esriSystem.IName)
    pDataLayer.Connect(pName)

    # Get service description, which includes the categories of wms layers
    pServiceDesc = pWMSGroupLayer.WMSServiceDescription

    return pWMSGroupLayer, pServiceDesc


def CreateWMSGroupLayer(url, outpath, visible, *layernames):
    # Luke Pinner 2014

    GetModule("esriCarto.olb")
    import comtypes.gen.esriCarto as esriCarto

    # Connect to WMS
    pWMSGroupLayer, pServiceDesc = ConnectWMS(url)

    # Find matching layers
    if layernames:
        pLayerDescs = [find_layer(pServiceDesc, layer) for layer in layernames]

        # Configure the layer before adding it to the map
        pWMSGroupLayer.Clear()
        for i, pLayerDesc in enumerate(pLayerDescs):
            if pLayerDesc is None:
                # raise  Exception('Unable to find layer "%s"'%layernames[i])
                # warnings.warn('Unable to find layer "%s"'%layernames[i])
                sys.stderr.write('Warning: Unable to find layer "%s"\n'%layernames[i])
                continue
            pWMSMapLayer = pWMSGroupLayer.CreateWMSLayer(pLayerDesc)
            CType(pWMSMapLayer, esriCarto.ILayer).Visible = visible
            pWMSGroupLayer.InsertLayer(pWMSMapLayer, i)

    pLayer = CType(pWMSGroupLayer, esriCarto.ILayer)
    pLayer.Visible = visible

    # Create a new LayerFile and populate it with the layer
    pLayerFile = NewObj(esriCarto.LayerFile, esriCarto.ILayerFile)
    pLayerFile.New(outpath)
    pLayerFile.ReplaceContents(pLayer)
    pLayerFile.Save()

    return outpath


# Luke Pinner 2012
def find_layer(servicedesc, layername):
    '''Recursive layer matching.
       Recursion is necessary to handle non data (folder/subfolder) "layers"'''
    layer = None
    for i in range(servicedesc.LayerDescriptionCount):
        layerdesc = servicedesc.LayerDescription(i)
        if layerdesc.LayerDescriptionCount == 0:
            if layerdesc.Name == layername:
                return layerdesc
            else:
                continue
        layer = find_layer(layerdesc, layername)
        if layer:
            break
    return layer


def MakeSelectionShapefile(input_shapefile, layer, output_shapefile):
    # Put in error trapping in case an error occurs when running tool
    try:

        # Make a layer from the feature class
        arcpy.MakeFeatureLayer_management(input_shapefile, layer)
        print("Made Feature Layer")

        # Within the layer select everything that is non grid code 255
        arcpy.SelectLayerByAttribute_management(
            layer, "NEW_SELECTION",
            "gridcode <> 255")

        print("Anything that is not grid code 255 selected")

        # Write the selected features to a new featureclass
        arcpy.CopyFeatures_management(layer, output_shapefile)
        print("New feature created with selected features" + OutLayerPoly)

        del input_shapefile
    except:
        print(arcpy.GetMessages())


def RasterToPolygon(output_shapefile2, input_raster):
    arcpy.RasterToPolygon_conversion(
        in_raster=input_raster,
        out_polygon_features=output_shapefile2,
        simplify="NO_SIMPLIFY",
        raster_field="Value",
        create_multipart_features="SINGLE_OUTER_PART",
        max_vertices_per_feature="")

    print("Raster converted to polygon, gridcode 255 is no value")

    del output_shapefile2


if __name__ == '__main__':

    arcpy.CreateTable_management("C:/TEMP", "cat_scrape.dbf")
    cat_scrape = "C:/TEMP/cat_scrape.dbf"
    arcpy.AddField_management(cat_scrape, "FILEID", "TEXT", field_length=100)
    arcpy.AddField_management(cat_scrape, "TYPE", "TEXT", field_length=25)
    arcpy.AddField_management(cat_scrape, "URL", "TEXT", field_length=256)
    arcpy.AddField_management(cat_scrape, "URLSERV", "TEXT", field_length=256)
    arcpy.DeleteField_management(cat_scrape, "Field1")
    """
    req = requests.request('GET','https://geocore.api.geo.ca/get')
    print(req)
    """

    # Luke Pinner 2012
    url = 'https://maps-cartes.services.geo.ca/server_serveur/services/VAC/military_memorials_en/MapServer/WMSServer?'
    # url = 'https://geo.statcan.gc.ca/geo_wa/services/NRN-RRN/NRN_WMS_EN/MapServer/WMSServer?'
    # url = 'https://maps.geogratis.gc.ca/wms/CBMT?'
    path1 = r'C:\Temp\test1.lyr'
    path2 = r'C:\Temp\test2.lyr'
    path3 = r'C:\Temp\test3.lyr'

    try:
        # All layers
        lyr = CreateWMSGroupLayer(url, path1, visible=True)
        # A couple of layers
        # lyr = CreateWMSGroupLayer(url, path2, False, 'lyr0', 'lyr1')
        # A single layer
        # lyr = CreateWMSGroupLayer(url, path3, False, 'lyr0')
        print("Create a layer file with WMS URL")
        print("WMS URL = " + url)
        print("Layer File = " + path1)

    except COMError as e:
        print('Unable to add WMS')

    # Creates an MXD
    #
    filepathMXD = 'c:/temp/testing123.mxd'
    CreateMXD(filepathMXD)

    # set MXD, set the dataframe to "Map" and add a layer to the bottom of the
    # dataframe, save the MXD and map as mxd2
    #
    mxd = arcpy.mapping.MapDocument(filepathMXD)
    df = arcpy.mapping.ListDataFrames(mxd, "Map")[0]
    addLayer = arcpy.mapping.Layer(path1)
    arcpy.mapping.AddLayer(df, addLayer, "BOTTOM")
    mxd.saveACopy(r"C:\TEMP\Project2.mxd")
    mxd2 = arcpy.mapping.MapDocument(r"C:\TEMP\Project2.mxd")
    del mxd, addLayer
    print("Layer Created in MXD")

    # Set dataframe and list layers in the mxd2
    data_frame = arcpy.mapping.ListDataFrames(mxd2, "Map")[0]
    layers = arcpy.mapping.ListLayers(mxd2)
    print(layers)

    # loop through the layers in the MXD and turn all layers visibility on
    for layer in layers:
        print(layer.longName)
        layer.visible = True
        # if layer.isGroupLayer:
        #   layer.visible = True
        # if layer.longName =="SubLayer Name":
        #   layer.visible = True
        # else:
        #   layer.visible = True
    
    print("WMS Layers visibility turned on in MXD")

    # save a copy of the mxd and map to mxd3
    mxd2.saveACopy(r"C:\TEMP\Project3.mxd")
    mxd3 = arcpy.mapping.MapDocument(r"C:\Temp\Project3.mxd")

    # loop through dataframes, and for Map return the spatial reference object
    # export a tiff from the dataframe with the width and height set in code
    for df in arcpy.mapping.ListDataFrames(mxd3):
        # df.rotation = 0
        SpatialReference = df.spatialReference
        print("Spatial Reference Object Created")

        # need to set coordinate system so output is correct
        outFile = r"C:\TEMP\\" + df.name + ".tif"
        arcpy.mapping.ExportToTIFF(
            mxd3, outFile,
            df,
            df_export_width=3200,
            df_export_height=1700,
            world_file=True)

    print("Output georeferenced TIFF created from WMS layer")

    del mxd2, mxd3
    arcpy.env.overwriteOutput = True

    # Replace a layer/table view name with a path to a dataset (which can be a
    # layer file) or create the layer/table view within the script
    # The following inputs are layers or table views: "Layers.tif"

    MapPoly = "C:/TEMP/Map_Poly.shp"
    Layers = "Layers"
    OutLayerPoly = "C:/TEMP/Map_Selection.shp"
    InRaster = "C:/TEMP/Map.tif"

    # Run raster to polygon transformation
    RasterToPolygon(MapPoly, InRaster)

    # make a selection from the shapefile based on everything that is not value
    # 255 and make a shapefile of that, for 255 in no value in WMS rasters
    MakeSelectionShapefile(MapPoly, Layers, OutLayerPoly)
    print("Layers = " + Layers)
    del MapPoly, Layers

    # StackExchange
    is_locked = list()

    for root, dirnames, filenames in os.walk('C:\TEMP'):
        for filename in filenames:
            full_filename = os.path.join(root, filename)
            try:
                pf = open(full_filename, 'a')
                pf.close()
            except:
                is_locked.append(full_filename)
                print("locked file = " + full_filename)

    finished_shapefile = "C:/TEMP/Map_Selection_Finished.shp"
    dissolved = "C:/TEMP/Map_Selection_Dissolve.shp"

    # add a field to the shapfile named "merge"
    arcpy.AddField_management(
        in_table=OutLayerPoly,
        field_name="MERGE",
        field_type="TEXT",
        field_precision="",
        field_scale="",
        field_length="5",
        field_alias="",
        field_is_nullable="NULLABLE",
        field_is_required="NON_REQUIRED", 
        field_domain="")

    print("Field Added")

    # calculate the field so that "merge" has a value of 1 for all polygons
    arcpy.CalculateField_management(
        in_table=OutLayerPoly,
        field="MERGE", expression="1",
        expression_type="VB",
        code_block="")

    print("Field Calculated")

    # dissolve based on the value 1 in "merge" creating one multipart polygon
    arcpy.Dissolve_management(
        in_features=OutLayerPoly,
        out_feature_class=dissolved,
        dissolve_field="MERGE",
        statistics_fields="",
        multi_part="MULTI_PART",
        unsplit_lines="DISSOLVE_LINES")

    print("Features Dissolved")

    del OutLayerPoly

    # take the dissolved multipart polygon and explode it into single part 
    # polygons
    arcpy.MultipartToSinglepart_management(
        in_features=dissolved,
        out_feature_class=finished_shapefile)

    print("Multi part to single part explosion")

    del dissolved

    # Add a field to count the vertices of polygons called "vertices"
    arcpy.AddField_management(
        in_table=finished_shapefile,
        field_name="VERTICES",
        field_type="FLOAT",
        field_precision="255",
        field_scale="0",
        field_length="",
        field_alias="",
        field_is_nullable="NULLABLE",
        field_is_required="NON_REQUIRED",
        field_domain="")

    print("Added field VERTICES")

    # calculate the field "vertices" with a count of vertices in each single
    # part polygon
    arcpy.CalculateField_management(
        finished_shapefile,
        "VERTICES",
        "!Shape!.pointCount-!Shape!.partCount",
        "PYTHON")

    print("Calculate the amount of vertices in VERTICES field")

    # Loop through the polygon shapefile and print a count of the number of
    # features/polygons
    PolygonCounter = 0
    with arcpy.da.SearchCursor(finished_shapefile, "MERGE") as cursor:
        for row in cursor:
            PolygonCounter = PolygonCounter + 1
    print("There are " + str(PolygonCounter) + " polygons")
    del row, cursor, PolygonCounter

    # set the projection of the shapefile based on the reference object from
    # earlier in the script
    arcpy.DefineProjection_management(finished_shapefile, SpatialReference)
    print("Projection defined for shapefile from original WMS")

    '''
    There will need to be a section for creating a shapefile and appending
    The polygons into it to ensure one common reference system: GCS WGS 1984
    '''

    # create a GeoJSON file from the shapefile which can be used to feed the
    # vertices/polygons into GeoCore
    arcpy.FeaturesToJSON_conversion(
        in_features=finished_shapefile,
        out_json_file="C:/TEMP/Map_Selection_Finished_Fe.json",
        format_json="FORMATTED", include_z_values="NO_Z_VALUES",
        include_m_values="NO_M_VALUES", geoJSON="GEOJSON")

    print("Create JSON output")

    del finished_shapefile

    del InRaster

    mxd3 = "C:/Temp/Project3.mxd"
    mxd2 = "C:/TEMP/Project2.mxd"
    mxd = 'C:/temp/testing123.mxd'
    path1 = 'C:/Temp/test1.lyr'
    InRaster = r"C:\TEMP\\" + df.name + ".tif"
    MapPoly2 = "C:/TEMP/Map_Poly.shp"
    OutLayerPoly = "C:/TEMP/Map_Selection.shp"
    dissolved = "C:/TEMP/Map_Selection_Dissolve.shp"
    finished_shapefile2 = "C:/TEMP/Map_Selection_Finished.shp"

    # Clean up directory, delete unnecessary files

    '''
    arcpy.Delete_management(mxd)
    print "Deleted mxd"
    arcpy.Delete_management(mxd2)
    print "Delete mxd2"
    arcpy.Delete_management(mxd3)
    print "Delete mxd3"
    arcpy.Delete_management(InRaster)
    print "Delete InRaster"
    arcpy.Delete_management(MapPoly2)
    print "Delete MapPoly"
    arcpy.Delete_management(OutLayerPoly)
    print "Delete OutLayerPoly"
    arcpy.Delete_management(dissolved)
    print "Delete dissolved"
    arcpy.Delete_management(finished_shapefile2)
    print "Delete finished_shapefile"
    arcpy.Delete_management(path1)
    print "Delete layer1"
    '''
    del mxd, mxd2, mxd3, InRaster, MapPoly2, OutLayerPoly, dissolved, finished_shapefile2, path1

sys.exit()
