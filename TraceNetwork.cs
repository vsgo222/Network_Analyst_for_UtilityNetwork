using System;
using System.Collections.Generic;
using System.Text;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.EditorExt;
using Miner.Interop;
using System.Data;
using System.Collections;
using System.Windows.Forms;
using ESRI.ArcGIS.Geometry;
using System.IO;
using System.Diagnostics;
using System.Data.OracleClient;

namespace RELGIS_DT_Load_Calculation
{
    enum CouplerPostion 
    { 
        NotPresent = 0, 
        On, 
        Off 
    }
    public class TraceNetwork
    {           
        #region Variable Declaration
        IApplication pApp;
        IMxDocument pMxDoc;
        IMap pMap;

        //FeatureLayers     
        #region Define FeatureLayers
        IFeatureLayer pSubStationFeatureLayer;        
        IFeatureLayer pCircuitBreakerFeatureLayer;
        IFeatureLayer pSwitchFeatureLayer;
        IFeatureLayer pBusBarFeatureLayer;
        IFeatureLayer pDistributedTransformerFeatureLayer;
        IFeatureLayer pLTCableFeatureLayer;
        IFeatureLayer pPillarFeatureLayer;
        IFeatureLayer pFuseFeatureLayer;
        IFeatureLayer pServicePointFeatureLayer;
        IFeatureLayer pTerminationPointFeatureLayer;
        IFeatureLayer pNetJunctionFeatureLayer;
        #endregion
        //TracedFeatureDetails DataTable
        #region Define DataTables
        System.Data.DataTable dtTracedResults = null;
        System.Data.DataTable dtErrorResults = null;
        System.Data.DataTable dtTappedLTCable = null;
        System.Data.DataTable dtPillarTable = null;
        System.Data.DataTable dtFUMPPillarTable = null;
        #endregion
        //Field Index Declaration
        #region Field Index Declarations
        private int iPillarLocation = -1;
        private int iTotalCktNos = -1;
        private int iFuseSwitchIDFldInd = -1;
        private int iFusePositionFldInd = -1;
        private int iDTRatingFldInd = -1;
        private int iDTSwitchIDFldInd = -1;
        private int iCOID = -1;
        private int iSwitchLinkIDFldInd = -1;
        private int iSwitchLinkPositionFldInd = -1;
        private int iFDRMGRNONTraceable = -1;
        private int iPillarSubType = -1;
        private int iErrorCount = 0;
        private int iCableSize = -1;        
        string sFieldsMissing = "";
        string sErrorMessage = "";
        int itesting = 0;
        #endregion
        IFeatureClass pLTCableFeatureClass = null;
        //Create Checklist for BusDropper
        #region Create Checklist for BusDropper
        ArrayList arrylstBusDropper = new ArrayList();
        ArrayList arrylstLTCableIDs = new ArrayList();
        ArrayList arrUniPillarID = new ArrayList();
        #endregion
        #region Create Checklist for FUMP type Fuse        
        ArrayList arrylstFuse = new ArrayList();
        ArrayList arrylstFUMPLTCableIDs = new ArrayList();
        ArrayList arrFUMPTypUniPillarID = new ArrayList();
        #endregion

        //CheckList for Pillar & BusDropper
        #region CheckList for Pillar & BusDropper
        ArrayList aryPillars = new ArrayList();
        ArrayList aryBD = new ArrayList();
        #endregion
        #endregion

        #region getting Arcmap layers & DataTable defination

        private System.Data.DataTable Create_DataTable(string _tablename)
        {            
            System.Data.DataTable oDataTable = new System.Data.DataTable(_tablename);
            //Create TracedEdge Details            
            oDataTable.Columns.Add("SrNo", System.Type.GetType("System.String"));
            oDataTable.Columns["SrNo"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("LTCableFeature", typeof(ESRI.ArcGIS.Geodatabase.IFeature));//Feature Type Details            
            oDataTable.Columns.Add("LTCableFeatureOID", System.Type.GetType("System.String"));
            oDataTable.Columns["LTCableFeatureOID"].DefaultValue = string.Empty;         
            oDataTable.Columns.Add("LTCableSize", System.Type.GetType("System.String"));//Get the Cross Section of current Cable
            oDataTable.Columns["LTCableSize"].DefaultValue = string.Empty;
            //This gives details of the first point of the Cable
            oDataTable.Columns.Add("End1OID", System.Type.GetType("System.String"));//Connected Object OID (Fuse/BusDropper)
            oDataTable.Columns["End1OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1Type", System.Type.GetType("System.String")); //Typeof Connected Feature: Fuse/BusDropper
            oDataTable.Columns["End1Type"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1No", System.Type.GetType("System.String"));//FuseNo
            oDataTable.Columns["End1No"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1Position", System.Type.GetType("System.String"));//Fuse: Normal Position (On/Off)
            oDataTable.Columns["End1Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1FacilityOID", System.Type.GetType("System.String")); //ObjectID of the Pillar
            oDataTable.Columns["End1FacilityOID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1FacilityType", System.Type.GetType("System.String")); // Type of Pillar
            oDataTable.Columns["End1FacilityType"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1FacilityName", System.Type.GetType("System.String"));//PillarNo
            oDataTable.Columns["End1FacilityName"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1BusBarLV", System.Type.GetType("System.String"));//BusBar Feature - LV
            oDataTable.Columns["End1BusBarLV"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1FacilityContainsBusCoupler", System.Type.GetType("System.String")); // Y/N
            oDataTable.Columns["End1FacilityContainsBusCoupler"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1BC1OID", System.Type.GetType("System.String"));//ObjectID of the First Connected BC to BusBar - LV (from Left)
            oDataTable.Columns["End1BC1OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1BC1Position", System.Type.GetType("System.String"));//Normal Position (On/Off)
            oDataTable.Columns["End1BC1Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1BC2OID", System.Type.GetType("System.String"));//ObjectID of the Second Connected BC to BusBar - LV (from Left)
            oDataTable.Columns["End1BC2OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1BC2Position", System.Type.GetType("System.String"));//Normal Position (On/Off)
            oDataTable.Columns["End1BC2Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("Remarks1", System.Type.GetType("System.String"));//Partial/Fully Loaded Pillar description
            oDataTable.Columns["Remarks1"].DefaultValue = string.Empty;
            //End2 Description
            oDataTable.Columns.Add("End2OID", System.Type.GetType("System.String"));//Connected Object OID (Fuse/BusDropper)
            oDataTable.Columns["End2OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2Type", System.Type.GetType("System.String")); //Typeof Connected Feature: Fuse/BusDropper
            oDataTable.Columns["End2Type"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2No", System.Type.GetType("System.String"));//FuseNo
            oDataTable.Columns["End2No"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2Position", System.Type.GetType("System.String"));//Fuse: Normal Position (On/Off)
            oDataTable.Columns["End2Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2FacilityOID", System.Type.GetType("System.String")); //ObjectID of the Pillar
            oDataTable.Columns["End2FacilityOID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2FacilityType", System.Type.GetType("System.String")); // Type of Pillar
            oDataTable.Columns["End2FacilityType"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2FacilityName", System.Type.GetType("System.String"));//PillarNo
            oDataTable.Columns["End2FacilityName"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2BusBarLV", System.Type.GetType("System.String"));//BusBar Feature - LV
            oDataTable.Columns["End2BusBarLV"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2FacilityContainsBusCoupler", System.Type.GetType("System.String")); // Y/N
            oDataTable.Columns["End2FacilityContainsBusCoupler"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2BC1OID", System.Type.GetType("System.String"));//ObjectID of the First Connected BC to BusBar - LV (from Left)
            oDataTable.Columns["End2BC1OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2BC1Position", System.Type.GetType("System.String"));//Normal Position (On/Off)
            oDataTable.Columns["End2BC1Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2BC2OID", System.Type.GetType("System.String"));//ObjectID of the Second Connected BC to BusBar - LV (from Left)
            oDataTable.Columns["End2BC2OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End2BC2Position", System.Type.GetType("System.String"));//Normal Position (On/Off)
            oDataTable.Columns["End2BC2Position"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("Remarks2", System.Type.GetType("System.String"));//Partial/Fully Loaded Pillar description
            oDataTable.Columns["Remarks2"].DefaultValue = string.Empty;
            // Details for connected ServicePoint
            oDataTable.Columns.Add("COID", System.Type.GetType("System.String"));//COID No. Connected to given SP
            oDataTable.Columns["COID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("UnitConsumption", System.Type.GetType("System.String"));//UnitConsumption consumed by given SP                        
            oDataTable.Columns["UnitConsumption"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("Flag", System.Type.GetType("System.String"));//UnitConsumption consumed by given SP                        
            oDataTable.Columns["Flag"].DefaultValue = "0";
            //Return
            return oDataTable;            
        }

        private System.Data.DataTable Create_ErrorTable(string _tablename)
        {
            System.Data.DataTable oDataTable = new System.Data.DataTable(_tablename);
            //Create TracedEdge Details            
            oDataTable.Columns.Add("SrNo", System.Type.GetType("System.String"));
            oDataTable.Columns["SrNo"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("LTCableFeatureOID", System.Type.GetType("System.String"));
            oDataTable.Columns["LTCableFeatureOID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("LayerName", System.Type.GetType("System.String"));
            oDataTable.Columns["LayerName"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("ErrorType", System.Type.GetType("System.String"));
            oDataTable.Columns["ErrorType"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("OID", System.Type.GetType("System.String"));
            oDataTable.Columns["OID"].DefaultValue = string.Empty;

            return oDataTable;
        }

        private System.Data.DataTable Create_TappedLTCableTable(string _tablename)
        {
            System.Data.DataTable oDataTable = new System.Data.DataTable(_tablename);
            //Create TracedEdge Details            
            oDataTable.Columns.Add("SrNo", System.Type.GetType("System.String"));
            oDataTable.Columns["SrNo"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("End1OID", System.Type.GetType("System.String"));
            oDataTable.Columns["End1OID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("CurrentFeatureOID", System.Type.GetType("System.String"));
            oDataTable.Columns["CurrentFeatureOID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("COID", System.Type.GetType("System.String"));
            oDataTable.Columns["COID"].DefaultValue = string.Empty;
            oDataTable.Columns.Add("UnitConsumption", System.Type.GetType("System.String"));
            oDataTable.Columns["UnitConsumption"].DefaultValue = "0";
            oDataTable.Columns.Add("Flag", System.Type.GetType("System.String"));
            oDataTable.Columns["Flag"].DefaultValue = "0";
            
            return oDataTable;
        }

        private System.Data.DataTable Create_PillarTable(string _tablename)
        {
            System.Data.DataTable oDataTable = new System.Data.DataTable(_tablename);
            try
            {
                oDataTable.Columns.Add("SrNo", System.Type.GetType("System.String"));
                oDataTable.Columns["SrNo"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("PillarOID", System.Type.GetType("System.String"));
                oDataTable.Columns["PillarOID"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-I", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-I"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-II", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-II"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-III", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-III"].DefaultValue = string.Empty;
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                
                #endregion
            }
            return oDataTable;
        }

        private System.Data.DataTable Create_FUMPType_PillarTable(string _tablename)
        {
            System.Data.DataTable oDataTable = new System.Data.DataTable(_tablename);
            try
            {
                oDataTable.Columns.Add("SrNo", System.Type.GetType("System.String"));
                oDataTable.Columns["SrNo"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("PillarOID", System.Type.GetType("System.String"));
                oDataTable.Columns["PillarOID"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-I", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-I"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-II", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-II"].DefaultValue = string.Empty;
                oDataTable.Columns.Add("BusBarLV-III", System.Type.GetType("System.String"));
                oDataTable.Columns["BusBarLV-III"].DefaultValue = string.Empty;
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region

                #endregion
            }
            return oDataTable;
        }
        
        private void getFeatureLayers()
        {            
            pCircuitBreakerFeatureLayer = Setting_Layers_Path.get_CircuitBreakerFeatureLayer(pMap);
            pDistributedTransformerFeatureLayer = Setting_Layers_Path.get_DistributedTransformerFeatureLayer(pMap);
            pSwitchFeatureLayer = Setting_Layers_Path.get_SwitchFeatureLayer(pMap);
            pFuseFeatureLayer = Setting_Layers_Path.get_FuseLayer(pMap);
            pBusBarFeatureLayer = Setting_Layers_Path.get_BusBarFeatureLayer(pMap);
            pServicePointFeatureLayer = Setting_Layers_Path.get_ServicePointFeatureLayer(pMap);
            pLTCableFeatureLayer = Setting_Layers_Path.get_LTCableFeatureLayer(pMap);
            pPillarFeatureLayer = Setting_Layers_Path.get_PillarFeatureLayer(pMap);
            pTerminationPointFeatureLayer = Setting_Layers_Path.get_TerminationPointFeatureLayer(pMap);
            pSubStationFeatureLayer = Setting_Layers_Path.get_SubStnFeatureLayer(pMap);
            pNetJunctionFeatureLayer = Setting_Layers_Path.get_NetJunctionFeatureLayer(pMap);
        }

        #endregion

        #region Setting Tracing Parameters

        public TraceNetwork(IApplication m_app)
        {
            try
            {
                if (m_app != null)
                {
                    //Get current map focus
                    pApp = m_app;
                    pMxDoc = pApp.Document as IMxDocument;
                    pMap = pMxDoc.FocusMap;
                    //Setting FeatureLayers
                    getFeatureLayers();
                    #region Create DataTables
                    dtTracedResults = Create_DataTable("LTCableTraceDetails");
                    dtErrorResults = Create_ErrorTable("ErrorList");
                    dtTappedLTCable = Create_TappedLTCableTable("TappedLTCable");
                    dtPillarTable = Create_PillarTable("PillarTable");
                    dtFUMPPillarTable = Create_FUMPType_PillarTable("FUMPType_PillarTable");
                    #endregion
                    #region Variable Declaration
                    iDTRatingFldInd = pDistributedTransformerFeatureLayer.FeatureClass.FindField("KVARATING");
                    if (iDTRatingFldInd < 0)
                        sFieldsMissing += " =>  KVA Rating - Distribution Transformer\n";

                    iDTSwitchIDFldInd = pDistributedTransformerFeatureLayer.FeatureClass.FindField("SWITCH_ID");
                    if (iDTSwitchIDFldInd < 0)
                        sFieldsMissing += " =>  Switch_ID - Distribution Transformer\n";

                    iSwitchLinkIDFldInd = pSwitchFeatureLayer.FeatureClass.FindField("SWITCH_ID");
                    if (iSwitchLinkIDFldInd < 0)
                        sFieldsMissing += " =>  Switch_ID - SwitchLink\n";

                    iSwitchLinkPositionFldInd = pSwitchFeatureLayer.FeatureClass.FindField("NORMALPOSITION_R");
                    if (iSwitchLinkPositionFldInd < 0)
                        sFieldsMissing += " =>  NORMALPOSITION_R - SwitchLink\n";

                    iFuseSwitchIDFldInd = pFuseFeatureLayer.FeatureClass.FindField("SWITCH_ID");
                    if (iFuseSwitchIDFldInd < 0)
                        sFieldsMissing += " =>  Switch ID - Fuse\n";

                    iFusePositionFldInd = pFuseFeatureLayer.FeatureClass.FindField("NORMALPOSITION_R");
                    if (iFusePositionFldInd < 0)
                        sFieldsMissing += " =>  Normal Position (A) - Fuse\n";

                    iPillarLocation = pPillarFeatureLayer.FeatureClass.FindField("PILLARNUMBER");
                    if (iPillarLocation < 0)
                        sFieldsMissing += " =>  Location of pillar - Pillar\n";

                    iTotalCktNos = pPillarFeatureLayer.FeatureClass.FindField("TOTALNUMBEROFCIRCUITS");
                    if (iTotalCktNos < 0)
                        sFieldsMissing += " =>  Total Number of Circuits - Pillar\n";

                    iCOID = pServicePointFeatureLayer.FeatureClass.FindField("CO_ID");
                    if (iCOID < 0)
                        sFieldsMissing += " =>  COID - Service point\n";

                    iFDRMGRNONTraceable = pBusBarFeatureLayer.FeatureClass.FindField("FDRMGRNONTRACEABLE");
                    if (iFDRMGRNONTraceable < 0)
                        sFieldsMissing += " =>  FDRMGRNONTRACEABLE - BusBar\n";
                    
                    iPillarSubType = pPillarFeatureLayer.FeatureClass.FindField("SUBTYPECD");
                    if (iPillarSubType < 0)
                        sFieldsMissing += " =>  SUBTYPECD - Pillar\n";

                    iCableSize = pLTCableFeatureLayer.FeatureClass.FindField("CABLESIZE");
                    if (iCableSize < 0)
                        sFieldsMissing += " =>  CABLESIZE - LTCable\n";
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void parameters(string sDT_SwitchID, IFeature pDistributeTransformerFeature_Selected, string sSubStn, string sDTRating)
        {
            try
            {
                pApp.StatusBar.ProgressBar.Message = "Initializing Parameters";
                //Parameters for Miner Miner Electric Trace Module
                UID oUID = new UIDClass();
                oUID.Value = "esriEdtiorExt.UtilityNetworkAnalysisExt";
                //Find extension
                INetworkAnalysisExt oNetworkAnalystExt = new UtilityNetworkAnalysisExtClass();
                oNetworkAnalystExt = (INetworkAnalysisExt)pApp.FindExtensionByCLSID(oUID);
                //set current geometric network for work
                IGeometricNetwork oCGN = oNetworkAnalystExt.CurrentNetwork;
                //Set elements to receive TracedFeature Collections
                IMMTracedElements tracedJunct;
                IMMTracedElements tracedEdg;
                //Tracing Function
                pApp.StatusBar.ProgressBar.Message = "";
                Tracing(oCGN, pDistributeTransformerFeature_Selected, out tracedJunct, out tracedEdg, sDT_SwitchID, sSubStn, sDTRating);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Tracing(IGeometricNetwork oCurrentGeometricNetwork, IFeature pSourceFeature, out IMMTracedElements tracedJunctions, out IMMTracedElements tracedEdges, string sDT_SwitchID, string sSubStn, string sDTRating)
        {
            pApp.StatusBar.ProgressBar.Message = "";
            tracedJunctions = new MMTracedElementsClass();
            tracedEdges = new MMTracedElementsClass();
            try
            {                          
                //Set Network Feature
                pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
                INetworkFeature pNWFeat = (INetworkFeature)pSourceFeature;
                ISimpleJunctionFeature pJuncFeat = (ISimpleJunctionFeature)pNWFeat;
                int CurrentStartEID = pJuncFeat.EID;
                IMMElectricTracing ElecTrace = new MMFeederTracerClass();
                IMMElectricTraceSettingsEx ElectricTraceSettings = new MMElectricTraceSettingsClass();
                //setting tracing conditions
                ElectricTraceSettings.UseFeederManagerCircuitSources = true;
                ElectricTraceSettings.UseFeederManagerProtectiveDevices = true;
                ElectricTraceSettings.RespectConductorPhasing = true;
                ElectricTraceSettings.RespectEnabledField = true;
                ElectricTraceSettings.RespectESRIBarriers = true;
                esriElementType StartElementType = esriElementType.esriETJunction;
                int[] barrJunc = new int[0];
                int[] barrEdges = new int[0];
                
                ElecTrace.TraceDownstream(oCurrentGeometricNetwork, ElectricTraceSettings, null, CurrentStartEID, 
                                                                    StartElementType, SetOfPhases.abc, 
                                                                    mmDirectionInfo.establishBySourceSearch, 0, 
                                                                    esriElementType.esriETEdge, barrJunc, barrEdges, 
                                                                    false, out tracedJunctions, out tracedEdges);
                try
                {
                    if (tracedEdges.Count > 0)
                    { 
                        Edges(tracedEdges, sDT_SwitchID, sSubStn, sDTRating);
                    }
                    else
                        MessageBox.Show("No TracedEdges");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
        }

        private void Edges(IMMTracedElements otracedEdges, string sDT_SwitchID, string sSubStn, string sDTRating)
        {
            pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
            itesting = 0;

            //Create DataTable

            //For each edge
            pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
            pLTCableFeatureClass = pLTCableFeatureLayer.FeatureClass; //returns FeatureClass of a FeatureLayer so defined
            int iLTCableFeatureClassID = pLTCableFeatureClass.FeatureClassID;//get the FeatureClassID
            IQueryFilter pQueryFilter = new QueryFilterClass();
            IMMTracedElement otracedEdge = otracedEdges.Next(); //Iterate through each TracedEdges so found
            //int iCount = 0;
            try
            {
                #region Create DataTable for Seggregation of LTCables
                DataTable dtSeggregatedLTCable = new DataTable();
                dtSeggregatedLTCable.Columns.Add("OID");
                dtSeggregatedLTCable.Columns.Add("FCID");
                dtSeggregatedLTCable.Columns.Add("EID");
                dtSeggregatedLTCable.Columns.Add("PreviousEID");
                dtSeggregatedLTCable.AcceptChanges();
                #endregion
                while (otracedEdge != null)
                {
                    pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
                    if (otracedEdge.FCID == iLTCableFeatureClassID)
                    {
                        #region Get ElementID Details
                        DataRow dRow = dtSeggregatedLTCable.NewRow();
                        dRow["OID"] = otracedEdge.OID.ToString();
                        dRow["FCID"] = otracedEdge.FCID.ToString();
                        dRow["EID"] = otracedEdge.EID.ToString();
                        dRow["PreviousEID"] = otracedEdge.PreviousEID.ToString();
                        dtSeggregatedLTCable.Rows.Add(dRow);
                        #endregion

                        try
                        {
                            pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
                            getTracedFeatureDetails(otracedEdge);
                        }
                        catch (Exception err)
                        {
                            MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
                        }
                    }
                    //Next Feature
                    otracedEdge = otracedEdges.Next();
                }

                #region Setting BusBar OID's for 2 BusCouplers

                SetBusBarLV4TwoBusCouplers();
                #region Setting BusBar OID's for 2 BusCouplers code modified at 31-May-10
                //foreach (DataRow dRowPillar in dtPillarTable.Rows)
                //{
                //    string sPillar = dRowPillar["PillarOID"].ToString();
                //    string sStrExpr = "End1FacilityOID = '" + sPillar + "'";
                //    DataRow[] dRows = dtTracedResults.Select(sStrExpr);
                //    if (dRows.Length > 0)
                //    {
                //        foreach (DataRow _dRows in dRows)
                //        {
                //            _dRows["End1BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                //            dtTracedResults.AcceptChanges();
                //        }
                //    }

                //    sStrExpr = "End2FacilityOID = '" + sPillar + "'";
                //    DataRow[] dRows1 = dtTracedResults.Select(sStrExpr);
                //    if (dRows.Length > 0)
                //    {
                //        foreach (DataRow _dRows in dRows1)
                //        {
                //            _dRows["End2BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                //            dtTracedResults.AcceptChanges();
                //        }
                //    }
                //}
                #endregion
                #endregion
                #region Setting BusBarLV OID's for 2 Fuse in FUMP
                SetBusBarLV4TwoFuse();
                #endregion
                ModifyTracedTable2includeTappedTableEntries();
                MessageBox.Show("Total Numbers of LTCables downstream :" + dtSeggregatedLTCable.Rows.Count.ToString());
                MessageBox.Show("Total No of Traced LTCables: " + itesting.ToString());                
                ArrangingEnd1End2Details();

                //Sequencing
                //GenerateSequence();

                #region ExportTables to Excel
                //If TracedTable does not contains any rows
                if (dtTracedResults.Rows.Count > 0)
                {
                    //MessageBox.Show(dtTracedResults.Rows.Count.ToString());
                    Setting_Layers_Path.Export2Excel(dtTracedResults, "RawTable", "RawTable");
                }
                else
                {
                    MessageBox.Show("No LTCables Traced");
                    Setting_Layers_Path.DeleteExcel(dtTracedResults, "RawTable");
                }

                //If SeggregatedTable does not contains any rows
                if (dtSeggregatedLTCable.Rows.Count > 0)
                { Setting_Layers_Path.Export2Excel(dtSeggregatedLTCable, "SeggregagedTable", "SeggregagedTable"); }
                else
                { Setting_Layers_Path.DeleteExcel(dtTracedResults, "SeggregagedTable"); }

                //If TappedCableTable does not contains any rows
                if (dtTappedLTCable.Rows.Count > 0)
                { Setting_Layers_Path.Export2Excel(dtTappedLTCable, "TappedCableTable", "TappedCableTable"); }
                else
                { Setting_Layers_Path.DeleteExcel(dtTappedLTCable, "TappedCableTable"); }

                //If ErrorTable does not contains any rows
                if (dtErrorResults.Rows.Count > 0)
                { Setting_Layers_Path.Export2Excel(dtErrorResults, "ErrorTable", "ErrorTable"); }
                else
                { Setting_Layers_Path.DeleteExcel(dtErrorResults, "ErrorTable"); }

                #endregion

                Hierrachy.HierrachyClass clsHierarchy = new Hierrachy.HierrachyClass(dtTracedResults, dtErrorResults);

                //Quote the Status of the Tool
                pApp.StatusBar.ProgressBar.Message = "Successfully Traced DT: " + sDT_SwitchID;
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
        }

        private void SetBusBarLV4TwoFuse()
        {
            try
            {
                foreach (DataRow dRowPillar in dtFUMPPillarTable.Rows)
                {
                    string sPillar = dRowPillar["PillarOID"].ToString();
                    string sStrExpr = "End1FacilityOID = '" + sPillar + "'";
                    DataRow[] dRows = dtTracedResults.Select(sStrExpr);
                    if (dRows.Length > 0)
                    {
                        foreach (DataRow _dRows in dRows)
                        {
                            _dRows["End1BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                            dtTracedResults.AcceptChanges();
                        }
                    }

                    sStrExpr = "End2FacilityOID = '" + sPillar + "'";
                    DataRow[] dRows1 = dtTracedResults.Select(sStrExpr);
                    if (dRows.Length > 0)
                    {
                        foreach (DataRow _dRows in dRows1)
                        {
                            _dRows["End2BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                            dtTracedResults.AcceptChanges();
                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
        }

        private void SetBusBarLV4TwoBusCouplers()
        {
            try
            {
                foreach (DataRow dRowPillar in dtPillarTable.Rows)
                {
                    string sPillar = dRowPillar["PillarOID"].ToString();
                    string sStrExpr = "End1FacilityOID = '" + sPillar + "'";
                    DataRow[] dRows = dtTracedResults.Select(sStrExpr);
                    if (dRows.Length > 0)
                    {
                        foreach (DataRow _dRows in dRows)
                        {
                            _dRows["End1BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                            dtTracedResults.AcceptChanges();
                        }
                    }

                    sStrExpr = "End2FacilityOID = '" + sPillar + "'";
                    DataRow[] dRows1 = dtTracedResults.Select(sStrExpr);
                    if (dRows.Length > 0)
                    {
                        foreach (DataRow _dRows in dRows1)
                        {
                            _dRows["End2BusBarLV"] = dRowPillar["BusBarLV-I"].ToString() + "_" + dRowPillar["BusBarLV-II"].ToString() + "_" + dRowPillar["BusBarLV-III"].ToString();
                            dtTracedResults.AcceptChanges();
                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
        }

        #region ModifyDataTables (i.e. Modify Traced Table 2 include Tapped Table Entries or ArrangingEnd1End2Details)
     
        private void ModifyTracedTable2includeTappedTableEntries()
        {
            #region ModifyTracedTable2includeTappedTableEntries main concept changes on 02-Jun-2010 due to include N-Level of tapping considerations
            //DataTable dtDistinctTappedCable = CommonFunc.SelectDistinct("DistinctTappedCables", dtTappedLTCable,
            //                                                                                            "End1OID");
            //foreach (DataRow dDistinctRow in dtDistinctTappedCable.Rows)
            //{
            //    double dUnitCon = 0.0;
            //    string sTemp = string.Empty;
            //    DataRow[] dRow = dtTappedLTCable.Select("End1OID = '" + dDistinctRow["End1OID"].ToString() + "'");
            //    DataRow[] dMainRow = dtTracedResults.Select("LTCableFeatureOID = '" + dDistinctRow["End1OID"].ToString() + "'");
            //    foreach (DataRow dtappedrow in dRow)
            //    {
            //        sTemp = dtappedrow["UnitConsumption"].ToString();
            //        if (sTemp != string.Empty)
            //            dUnitCon += Convert.ToDouble(dtappedrow["UnitConsumption"].ToString());
            //        else
            //            dUnitCon += 0.00;

            //        sTemp = string.Empty;
            //    }
            //    if (dMainRow.Length > 0)
            //    {
            //        sTemp = dMainRow[0]["UnitConsumption"].ToString();
            //        dUnitCon += Convert.ToDouble(sTemp);
            //        dMainRow[0]["UnitConsumption"] = dUnitCon.ToString();
            //        dtTracedResults.AcceptChanges();
            //    }
            //}
            #endregion

            #region New conde to include N-Level Tapping conditions
            DataTable dtDistinctTappedCable = CommonFunc.SelectDistinct("DistinctTappedCables", dtTappedLTCable,
                                                                                                        "End1OID");
            foreach (DataRow dRow in dtTracedResults.Rows)
            {
                string strOID = dRow["LTCableFeatureOID"].ToString();
                DataRow[] dSelectedRow = dtDistinctTappedCable.Select("End1OID = '" + strOID + "'");
                if (dSelectedRow.Length != 0)
                {
                    //Get Details of Tapped Cable
                    foreach (DataRow _dSelectedRow in dSelectedRow)
                    {
                        double dUnitCon = getLoad4TappedCable(strOID);
                        dRow["UnitConsumption"] = dUnitCon.ToString();
                        dtTracedResults.AcceptChanges();
                    }
                }
            }
            #endregion
        }

        private double getLoad4TappedCable(string _strOID)
        {
            double dUnitCon = 0.0;
            string sTemp = string.Empty;
            try
            {
                DataRow[] dRows = dtTappedLTCable.Select("End1OID = '" + _strOID + "'");
                if (dRows.Length != 0)
                {
                    foreach (DataRow dRow in dRows)
                    {
                        if(dRow["Flag"].ToString() != "1")
                        {
                            sTemp = dRow["UnitConsumption"].ToString();
                            string strOID = dRow["End1OID"].ToString();
                            if (sTemp != string.Empty || sTemp != "" || sTemp != " ")
                                dUnitCon += Convert.ToDouble(dRow["UnitConsumption"].ToString());
                            else
                                dUnitCon += 0.00;

                            dRow["Flag"] = "1";
                            dtTappedLTCable.AcceptChanges();

                            sTemp = string.Empty;
                            dUnitCon += getLoad4TappedCable(strOID);                            
                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
            return dUnitCon;
        }

        private void ArrangingEnd1End2Details()
        {
            try
            {
                string sStrExpr = "(End1FacilityOID = 'SP' AND End1FacilityType = 'SP' AND End1FacilityName = 'SP')";
                DataRow[] dRows = dtTracedResults.Select(sStrExpr);
                if (dRows.Length > 0)
                {

                    foreach (DataRow _dRows in dRows)
                    {
                        object[] arrList = _dRows.ItemArray;
                        int j = 4;
                        for (int i = 18; i < 31; i++)
                        {                            
                            _dRows[j] = arrList[i].ToString();
                            j++;                            
                        }
                        j = 18;
                        for (int i = 4; i < 17; i++)
                        {
                            _dRows[j] = arrList[i].ToString();
                            j++;
                        }                        
                        dtTracedResults.AcceptChanges();
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
            }
        }

        #endregion ModifyDataTables

        #endregion

        #region Main Loop (Check conditions involved for each LTCable)

        private void getTracedFeatureDetails(IMMTracedElement otracedEdge)
        {
            
            IFeature pFeature = null;
            pApp.StatusBar.ProgressBar.Message = "Tracing Features...";
            pFeature = pLTCableFeatureClass.GetFeature(otracedEdge.OID);//get the Object ID of the <--> Feature          
            ArrayList arryFeatureOID= new ArrayList();           
            try
            {
                if (pFeature != null)
                {  
                    //Create New Row to add details of current feature
                    DataRow oNewRow = dtTracedResults.NewRow();
                    oNewRow["LTCableFeature"] = pFeature;
                    oNewRow["LTCableFeatureOID"] = pFeature.OID.ToString();
                    oNewRow["LTCableSize"] = pFeature.get_Value(iCableSize).ToString();
                    //Find whether the Cable is connected to DT
                    bool bConnected2DTCable = get_ConnectedDTDetails(ref oNewRow, pFeature);
                    #region If Cable is connected to DT
                    if (bConnected2DTCable)
                    {
                        pApp.StatusBar.ProgressBar.Message = "Finding Duplicate Features, if any";
                        //if (!((FindForDuplicateCable(dtTracedResults, oNewRow))))
                        if (!(arrylstLTCableIDs.Contains(pFeature.OID.ToString())))
                        {
                            pApp.StatusBar.ProgressBar.Message = "";
                            dtTracedResults.Rows.Add(oNewRow);
                            arrylstLTCableIDs.Add(pFeature.OID.ToString());
                            itesting++;
                        }
                    }
                    #endregion
                    #region If Cable is not connected to DT
                    else
                    {
                        /// <summary>
                        /// Since LTCable is not connected to DT we need to split the LTCable into two parts
                        /// First, would be the first point of the LTCable i.e. assuming that digitization is done in right way and not on the reverse way
                        /// Second, point would be the last point of the LTCable
                        /// </summary>
                        //Convert the LTCable to polyline
                        IPolyline pPolyLine = (IPolyline)pFeature.Shape;
                        //Get the point collection within the polyline
                        IPointCollection pPointCollection = (IPointCollection)pPolyLine;
                        if (oNewRow["LTCableFeatureOID"].ToString() == "288562" || oNewRow["LTCableFeatureOID"].ToString() == "288558")
                        {
                        }
                        if (pPointCollection.PointCount > 0)
                        {
                            //Get the First point of LTCable
                            IPoint pFirstPoint = pPointCollection.get_Point(0);
                            bool bSnappingErrorCheck = false;
                            bool bFirstConnectedFeature = getEnd1ConnectedFeatures(pFirstPoint, pFeature, ref oNewRow,
                                                                                                    ref bSnappingErrorCheck);
                            bool bSecondConnectedFeature = false; IPoint pLastPoint = null;
                            //Get the Last point of LTCable
                            if (bSnappingErrorCheck == false)
                            {
                                pLastPoint = pPointCollection.get_Point(pPointCollection.PointCount - 1);
                                bSecondConnectedFeature = getEnd2ConnectedFeature(pLastPoint, pFeature, ref oNewRow,
                                                                                                    ref bSnappingErrorCheck);
                            }
                            if (bFirstConnectedFeature && bSecondConnectedFeature)
                            {                               
                                pApp.StatusBar.ProgressBar.Message = "Finding Duplicate Features, if any";
                                //if (!((FindForDuplicateCable(dtTracedResults, oNewRow))))
                                if (!(arrylstLTCableIDs.Contains(pFeature.OID.ToString())))
                                {
                                    pApp.StatusBar.ProgressBar.Message = "";
                                    dtTracedResults.Rows.Add(oNewRow);
                                    arrylstLTCableIDs.Add(pFeature.OID.ToString());
                                    itesting++;
                                }
                            }
                            else
                            {
                                //MessageBox.Show(pFeature.OID.ToString());
                                if (bSnappingErrorCheck == false)
                                {
                                    //Find whether LTCable is Tapped Cable
                                    DataRow oTappedRow = dtTappedLTCable.NewRow();
                                    bool bTappedConnectedFeature = false;
                                    bTappedConnectedFeature = GetTappedCableDetails(pFirstPoint, pLastPoint, pFeature, ref oTappedRow);
                                    if (bTappedConnectedFeature)
                                    {
                                        pApp.StatusBar.ProgressBar.Message = "Finding Duplicate Features, if any";
                                        if (!(arrylstLTCableIDs.Contains(pFeature.OID.ToString())))
                                        {
                                            pApp.StatusBar.ProgressBar.Message = "";
                                            dtTappedLTCable.Rows.Add(oTappedRow);
                                            arrylstLTCableIDs.Add(pFeature.OID.ToString());
                                            itesting++;
                                            //MessageBox.Show(pFeature.OID.ToString());
                                        }
                                    }
                                    //Find whether LTCable is Tapped but not connected to SP
                                    else
                                    {
                                        bool bConnectedLTCableWithoutTapping = false;
                                        DataRow oConnectedLTCableWithoutTappingRow = dtTracedResults.NewRow();
                                        bConnectedLTCableWithoutTapping = GetLTCableWithoutTapping(pFirstPoint,
                                                                                            pLastPoint, pFeature,
                                                                            ref oConnectedLTCableWithoutTappingRow);
                                        if (bConnectedLTCableWithoutTapping)
                                        {
                                            pApp.StatusBar.ProgressBar.Message = "Finding Duplicate Features, if any";
                                            if (!(arrylstLTCableIDs.Contains(pFeature.OID.ToString())))
                                            {
                                                pApp.StatusBar.ProgressBar.Message = "";
                                                dtTracedResults.Rows.Add(oConnectedLTCableWithoutTappingRow);
                                                arrylstLTCableIDs.Add(pFeature.OID.ToString());
                                                itesting++;
                                                //MessageBox.Show(pFeature.OID.ToString());
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                  //  MessageBox.Show(pFeature.OID.ToString());
                                }
                            }
                        }
                    }
                    #endregion If Cable is not connected to DT
                }
                else
                {
                    FillErrorTable("No Feature found", "LTCable", otracedEdge.OID.ToString());                               
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
        }

        #endregion Main Loop (Check conditions involved for each LTCable)

        #region Excel Funtions

        private void ExportToExcel(System.Data.DataTable _datatable, string _str)
        {           
            string FilePath = @"C:\ExcelData\" + _str.ToString() + ".xls";
            if (File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }
            // Excel object references.
            Excel.Application m_objExcel = null;
            Excel.Workbooks m_objBooks = null;
            Excel._Workbook m_objBook = null;
            Excel.Sheets m_objSheets = null;
            Excel._Worksheet m_objSheet = null;
            Excel.Range m_objRange = null;
            Excel.Font m_objFont = null;
            // Frequenty-used variable for optional arguments.
            object m_objOpt = System.Reflection.Missing.Value;
            // Paths used by the sample code for accessing and storing data.
            object m_strSampleFolder = "C:\\ExcelData\\";
            //string m_strNorthwind = "C:\\Program Files\\Microsoft Office\\Office10\\Samples\\Northwind.mdb";

            // Start a new workbook in Excel.
            m_objExcel = new Excel.Application();
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            // Add data to cells of the first worksheet in the new workbook.

            #region Set Column Headers 
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
            m_objRange = m_objSheet.get_Range("A1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "LTCableFeature");
            m_objRange = m_objSheet.get_Range("B1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "LTCableFeatureOID");

            m_objRange = m_objSheet.get_Range("D1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1OID");
             m_objRange = m_objSheet.get_Range("E1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1Type");
            m_objRange = m_objSheet.get_Range("F1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1No");
            m_objRange = m_objSheet.get_Range("G1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1Position");
              m_objRange = m_objSheet.get_Range("H1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1FacilityOID");
             m_objRange = m_objSheet.get_Range("I1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1FacilityType");
            m_objRange = m_objSheet.get_Range("J1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1FacilityName");
            m_objRange = m_objSheet.get_Range("K1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1BusBarLV");

            m_objRange = m_objSheet.get_Range("L1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1FacilityContainsBusCoupler");
             m_objRange = m_objSheet.get_Range("M1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1BC1OID");
            m_objRange = m_objSheet.get_Range("N1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1BC1Position");
            m_objRange = m_objSheet.get_Range("O1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1BC2OID");
              m_objRange = m_objSheet.get_Range("P1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End1BC2Position");
            m_objRange = m_objSheet.get_Range("Q1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "Remarks2"); 
            //End 2 Details
             m_objRange = m_objSheet.get_Range("R1", m_objOpt);           
            m_objRange.set_Value(m_objOpt, "End2OID");
             m_objRange = m_objSheet.get_Range("S1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2Type");
            m_objRange = m_objSheet.get_Range("T1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2No");
            m_objRange = m_objSheet.get_Range("U1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2Position");
              m_objRange = m_objSheet.get_Range("V1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2FacilityOID");
             m_objRange = m_objSheet.get_Range("W1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2FacilityType");
            m_objRange = m_objSheet.get_Range("X1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2FacilityName");
            m_objRange = m_objSheet.get_Range("Y1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2BusBarLV");

            m_objRange = m_objSheet.get_Range("AA1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2FacilityContainsBusCoupler");
             m_objRange = m_objSheet.get_Range("AB1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2BC1OID");
            m_objRange = m_objSheet.get_Range("AC1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2BC1Position");
            m_objRange = m_objSheet.get_Range("AD1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2BC2OID");
              m_objRange = m_objSheet.get_Range("AE1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "End2BC2Position");
            m_objRange = m_objSheet.get_Range("AF1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "Remarks2");  
            //COID Details                            
            m_objRange = m_objSheet.get_Range("AH1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "COID");
            m_objRange = m_objSheet.get_Range("AI1", m_objOpt);
            m_objRange.set_Value(m_objOpt, "UnitConsumption");
            #endregion

            int i = 2;

            #region Input data into Excel File
            foreach (DataRow pRow in _datatable.Rows)
            {
                //Add data to the excel table
                m_objRange = m_objSheet.get_Range("A" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["LTCableFeature"].ToString());
                m_objRange = m_objSheet.get_Range("B" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["LTCableFeatureOID"].ToString());

                m_objRange = m_objSheet.get_Range("D" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1OID"].ToString());
                m_objRange = m_objSheet.get_Range("E" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1Type"].ToString());
                m_objRange = m_objSheet.get_Range("F" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1No"].ToString());
                m_objRange = m_objSheet.get_Range("G" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1Position"].ToString());
                m_objRange = m_objSheet.get_Range("H" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1FacilityOID"].ToString());
                m_objRange = m_objSheet.get_Range("I" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1FacilityType"].ToString());
                m_objRange = m_objSheet.get_Range("J" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1FacilityName"].ToString());
                m_objRange = m_objSheet.get_Range("K" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1BusBarLV"].ToString());

                m_objRange = m_objSheet.get_Range("L" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1FacilityContainsBusCoupler"].ToString());
                m_objRange = m_objSheet.get_Range("M" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1BC1OID"].ToString());
                m_objRange = m_objSheet.get_Range("N" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1BC1Position"].ToString());
                m_objRange = m_objSheet.get_Range("O" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1BC2OID"].ToString());
                m_objRange = m_objSheet.get_Range("P" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End1BC2Position"].ToString());
                m_objRange = m_objSheet.get_Range("Q" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["Remarks1"].ToString()); 

                //End 2 Details
                m_objRange = m_objSheet.get_Range("R" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2OID"].ToString());
                m_objRange = m_objSheet.get_Range("S" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2Type"].ToString());
                m_objRange = m_objSheet.get_Range("T" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2No"].ToString());
                m_objRange = m_objSheet.get_Range("U" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2Position"].ToString());
                m_objRange = m_objSheet.get_Range("V" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2FacilityOID"].ToString());
                m_objRange = m_objSheet.get_Range("W" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2FacilityType"].ToString());
                m_objRange = m_objSheet.get_Range("X" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2FacilityName"].ToString());
                m_objRange = m_objSheet.get_Range("Y" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2BusBarLV"].ToString());
                m_objRange = m_objSheet.get_Range("AA" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2FacilityContainsBusCoupler"].ToString());
                m_objRange = m_objSheet.get_Range("AB" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2BC1OID"].ToString());
                m_objRange = m_objSheet.get_Range("AC" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2BC1Position"].ToString());
                m_objRange = m_objSheet.get_Range("AD" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2BC2OID"].ToString());
                m_objRange = m_objSheet.get_Range("AE" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["End2BC2Position"].ToString());
                m_objRange = m_objSheet.get_Range("AF" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["Remarks2"].ToString()); 
                //COID Details                            
                m_objRange = m_objSheet.get_Range("AH" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["COID"].ToString());                
                m_objRange = m_objSheet.get_Range("AI" + i, m_objOpt);
                m_objRange.set_Value(m_objOpt, pRow["UnitConsumption"].ToString());
                                
                i++;
            }

            #endregion

            // Apply bold to cells A1:B1.
            m_objRange = m_objSheet.get_Range("A1", "AG1");
            m_objFont = m_objRange.Font;
            m_objFont.Bold = true;

            // Save the workbook and quit Excel.
            m_objBook.SaveAs(m_strSampleFolder + (_str.ToString() + ".xls"), m_objOpt, m_objOpt,
                                                m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                                                                m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
            m_objBook.Close(false, m_objOpt, m_objOpt);
            m_objExcel.Quit();
            m_objExcel = null;
            killExcel();
        }

        private void killExcel()
        {
            try
            {
                Process[] ps = Process.GetProcesses();
                foreach (Process p in ps)
                {
                    if (p.ProcessName.ToLower().Equals("excel"))
                    {
                        p.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        #endregion

        #region FindFeatures

        private bool get_ConnectedDTDetails(ref DataRow pRow, IFeature _LTEdge)
        {
            bool bConnected2DTCable = false;

            IFeature pDTFeature = CommonGISFunc.FirstSpatialQueryFeature(_LTEdge.Shape, pDistributedTransformerFeatureLayer, esriSpatialRelEnum.esriSpatialRelIntersects, "");
            //IFeature pEndFeature = null; IFeature pPillarFeature = null;
            if (pDTFeature != null)
            {
               // itesting++;
                pRow["End1OID"] = pDTFeature.OID.ToString();
                pRow["End1Type"] = pDTFeature.Class.ObjectClassID.ToString();
                pRow["End1No"] = pDTFeature.get_Value(iDTSwitchIDFldInd).ToString() + "_" + pDTFeature.get_Value(pDTFeature.Fields.FindField("KVARATING")).ToString();
                pRow["End1Position"] = "1";
                pRow["End1FacilityOID"] = "DT";
                pRow["End1FacilityType"] = "DT";
                pRow["End1FacilityName"] = "DT";
                pRow["End1BusBarLV"] = "";
                pRow["End1FacilityContainsBusCoupler"] = "N";
                pRow["End1BC1OID"] = "";
                pRow["End1BC1Position"] = "";
                pRow["End1BC2OID"] = "";
                pRow["End1BC2Position"] = "";

                //Get Details of other end point
                getEnd2DTDetails(ref pRow, _LTEdge, ref bConnected2DTCable);
                //if (b2ndEndDetails)
                //{
                bConnected2DTCable = true;
                //}
            }
            else
            {
                //error table entry for no BusBarLV
               // FillErrorTable("No DT found connected to LTCable", "DT / LTCable", _LTEdge.OID.ToString());              
              //  bConnected2DTCable = false;
            }
            return bConnected2DTCable;
        }

        private void getEnd2DTDetails(ref DataRow pRow, IFeature _LTEdge,ref bool bConnected2DTCable)
        {
            //bool b2ndEndDetails = false;
            ArrayList arrBDList = new ArrayList();
            IFeature pBDFeature = CommonGISFunc.FirstSpatialQueryFeature(_LTEdge.Shape,
                                                    pBusBarFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "SUBTYPECD = 4"); //Get BusDropper LV
            if (pBDFeature != null)
            {
                //b2ndEndDetails = true;
                arrBDList.Add(pBDFeature.OID);
                IFeature pFuseFeature = CommonGISFunc.FirstSpatialQueryFeature(pBDFeature.Shape,
                                                                    pFuseFeatureLayer,
                                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "");
                if (pFuseFeature != null)
                {
                    //Get the Fuse Details
                    pRow["End2OID"] = pFuseFeature.OID.ToString();
                    pRow["End2Type"] = pFuseFeature.Class.ObjectClassID.ToString();
                    pRow["End2No"] = pFuseFeature.get_Value(iFuseSwitchIDFldInd).ToString();
                    pRow["End2Position"] = pFuseFeature.get_Value(iFusePositionFldInd).ToString();
                    pRow["COID"] = "0";
                    pRow["UnitConsumption"] = "0";
                    getPillarFeature(pFuseFeature, ref pRow, 2);
                    GenericCollection pBusDroperFeatureColl = new GenericCollection();
                    CommonGISFunc.AllSpatialQueryFeatures(pFuseFeature.Shape, pBusBarFeatureLayer,
                                                                        esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                        "SUBTYPECD = 4",
                                                                        ref pBusDroperFeatureColl);
                    if (pBusDroperFeatureColl.Count > 0)
                    {
                        foreach (IFeature pBD in pBusDroperFeatureColl)
                        {
                            if ((!arrBDList.Contains(pBD.OID)) && (pBD.get_Value(iFDRMGRNONTraceable).ToString() == "0"))
                            {
                                IFeature pBusBarLV = CommonGISFunc.FirstSpatialQueryFeature(pBD.Shape,
                                                    pBusBarFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "SUBTYPECD = 3");
                                if (pBusBarLV != null)
                                {
                                    getPillarDetails(pBusBarLV);
                                    pRow["End2BusBarLV"] = pBusBarLV.OID.ToString();
                                    IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                                    IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                                    IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                                    IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                                    BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 2, pBusBarLV);
                                }
                            }
                        }
                    }
                }
                else
                {
                    FillErrorTable("No Fuse found connected to BusDropper", "BusDropper / Fuse", pBDFeature.OID.ToString());
                    bConnected2DTCable = false;
                }
            }
            else
            {
                //error table entry for no BusBarLV
                iErrorCount++;
                DataRow pErrorRow = dtErrorResults.NewRow();
                pErrorRow["SrNo"] = iErrorCount.ToString();
                pErrorRow["LayerName"] = "LTCable / BusDropper";
                pErrorRow["ErrorType"] = "No BusDropper found connected to LTCable";
                pErrorRow["OID"] = _LTEdge.OID.ToString();
                dtErrorResults.Rows.Add(pErrorRow);
                bConnected2DTCable = false;
            }
        }

        private bool getEnd1ConnectedFeatures(IPoint _pPoint, IFeature pCurrentLTCableFeature ,ref DataRow pRow, ref bool _bSnappingErrorCheck)
        {
            bool bGetConnectedFeature = false;
            if (pCurrentLTCableFeature.OID.ToString() == "2300955")
            {
            }
            //If connected Feature is TerminationPoint
            #region If feature is connected to TerminationPoint commented
            //IFeature pTerminationPoint = getIntersectingFeature(_pPoint, pTerminationPointFeatureLayer, "");
            //if (pTerminationPoint != null)
            //{
            //    pRow["End1OID"] = pTerminationPoint.OID.ToString();
            //    pRow["End1Type"] = pTerminationPoint.FeatureType.ToString();
            //    pRow["End1No"] = pTerminationPoint.OID.ToString();
            //    pRow["End1Position"] = "1";
            //    IFeature pPillarFeature = getIntersectingFeature(pTerminationPoint, pPillarFeatureLayer, "");
            //    if (pPillarFeature != null)
            //    {
            //        pRow["End1FacilityOID"] = pPillarFeature.OID.ToString();
            //        pRow["End1FacilityType"] = "MP";
            //        pRow["End1FacilityName"] = pPillarFeature.get_Value(iPillarLocation);
            //    }
            //    //If the connected feature is a busdropper
            //    IFeature pBDFeature = getIntersectingFeature(pTerminationPoint, pBusBarFeatureLayer, "SUBTYPECD = 4");
            //    if (pBDFeature != null)
            //    {
            //        IFeature pBusBarLV = getIntersectingFeature(pBDFeature, pBusBarFeatureLayer, "SUBTYPECD = 3");
            //        if (pBusBarLV != null)
            //        {
            //            pRow["End1BusBarLV"] = pBusBarLV.OID.ToString();
            //            pRow["End1FacilityContainsBusCoupler"] = "N";
            //            pRow["End1BC1OID"] = "0";
            //            pRow["End1BC1Position"] = "0";
            //            pRow["End1BC2OID"] = "0";
            //            pRow["End1BC2Position"] = "0";
            //            bGetConnectedFeature = true;
            //        }
            //        else
            //        {
            //            //error table entry for no BusBarLV
            //            iErrorCount++;
            //            DataRow pErrorRow = dtErrorResults.NewRow();
            //            pErrorRow["SrNo"] = iErrorCount.ToString();
            //            pErrorRow["LayerName"] = "BusDropper / BusBarLV";
            //            pErrorRow["ErrorType"] = "No BusBarLV found connected to BusDropper";
            //            pErrorRow["OID"] = pBDFeature.OID.ToString();
            //            dtErrorResults.Rows.Add(pErrorRow);
            //        }
            //    }
            //    else
            //    {
            //        //error table entry for no BDFeature 
            //        iErrorCount++;
            //        DataRow pErrorRow = dtErrorResults.NewRow();
            //        pErrorRow["SrNo"] = iErrorCount.ToString();
            //        pErrorRow["LayerName"] = "Termination / BusDropper";
            //        pErrorRow["ErrorType"] = "No BusDropper found connected to Termination Point";
            //        pErrorRow["OID"] = pTerminationPoint.OID.ToString();
            //        dtErrorResults.Rows.Add(pErrorRow);
            //    }
            //}
            #endregion
            //If connected Feature is BusDropper
            //IFeature pBDFeature = getIntersectingFeature(_pPoint, pBusBarFeatureLayer, "SUBTYPECD = 4");
            IFeature pBDFeature = Setting_Layers_Path.GetFirstBufferedFeature((IGeometry)_pPoint, 0.00101,
                                                                             pBusBarFeatureLayer, "SUBTYPECD = 4");
            if (pBDFeature != null)
            {
                IFeature pFuse = getIntersectingFeature(pBDFeature, pFuseFeatureLayer, "");
                #region Chking & getting Fuse details
                if (pFuse != null)
                {
                    pRow["End1OID"] = pFuse.OID.ToString();
                    pRow["End1Type"] = pFuse.Class.ObjectClassID.ToString();
                    pRow["End1No"] = pFuse.OID.ToString();
                    pRow["End1Position"] = pFuse.get_Value(iFusePositionFldInd).ToString();

                    getPillarFeature(pFuse, ref pRow, 1);
                    //pBDFeature = getBusDropperFeature(pFuse, pBDFeature.OID.ToString());
                    pBDFeature = getFeature4GenericCollection(pFuse, pBusBarFeatureLayer ,pBDFeature.OID.ToString());
                    #region Chking & getting BusCuopler details
                    if (pBDFeature != null)
                    {
                        IFeature pBusBarLV = getIntersectingFeature(pBDFeature, pBusBarFeatureLayer,
                                                                                            "SUBTYPECD = 3");
                        if (pBusBarLV != null)
                        {
                            pRow["End1BusBarLV"] = pBusBarLV.OID.ToString();
                            IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                            IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                            IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                            IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                            BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 1, pBusBarLV);
                            bGetConnectedFeature = true;
                        }
                        else
                        {
                            //error table entry for no BusBarLV 
                            FillErrorTable("No Feature found connected to BusDropper", "BusDropper", pBDFeature.OID.ToString());
                            _bSnappingErrorCheck = true;                          
                        }
                    }
                    else
                    {
                        //error table entry for no BDFeature 
                        FillErrorTable("No Feature found connected to Fuse", "Fuse", pFuse.OID.ToString());                       
                        _bSnappingErrorCheck = true;
                    }
                    #endregion Chking & getting BusCuopler details
                }
                else //If no Fuse is found
                {
                    //Check if the BusDropper is connected to BusBar-LV 
                    #region Find BusBar-LV connected to BusDropper
                    IFeature pBusBarLV = getIntersectingFeature(pBDFeature, pBusBarFeatureLayer,
                                                                                            "SUBTYPECD = 3");
                    if (pBusBarLV != null)
                    {
                        pRow["End1BusBarLV"] = pBusBarLV.OID.ToString();
                        IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                        IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                        IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                        IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                        BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 1, pBusBarLV);
                        bGetConnectedFeature = true;
                        getPillarFeature(pBusBarLV, ref pRow, 1);
                        //fill Fuse property columns with BusDropper properties 
                        pRow["End1OID"] = pBDFeature.OID.ToString();
                        pRow["End1Type"] = pBDFeature.Class.ObjectClassID.ToString();
                        pRow["End1No"] = pBDFeature.OID.ToString();
                        pRow["End1Position"] = "1"; //Assuming that BusDropper can never be OFF
                    }
                    else
                    {
                        //error table entry for no BusBarLV 
                        FillErrorTable("No Feature found connected to BusDropper", "BusDropper", pBDFeature.OID.ToString());                      
                        _bSnappingErrorCheck = true;
                    }
                    #endregion
                }
                #endregion Chking & getting Fuse details
            }
            else
            {
                //Find connected BusBar-LV
                #region Find connected BusBar-LV
                IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(_pPoint, 0.014,
                                                                            pBusBarFeatureLayer, "SUBTYPECD = 3");
                if (pBusBarLV != null)
                {
                    IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBusBarLV.Shape,
                                                                            pPillarFeatureLayer,
                                                                            esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                            "");
                   int iFUMPChk = 0;
                    if (pPillarFeature != null)
                    {

                        string strSubType = CommonGISFunc.GetFieldValueFN(pPillarFeature, "SUBTYPECD", true);
                        if (strSubType == "4")
                        {
                            iFUMPChk = 1;
                        }
                    }
                    if (iFUMPChk == 1)
                    {
                        GetFUMPDetails(pBusBarLV, ref pRow, 1);
                    }
                    else
                    {                  
                        #region MiniPillar with BusBar-LV connected directly to LTCable
                        /// <Explanation>
                        /// If BusBar is not connected to Fuse
                        /// i.e Pillar is not type of FUMP
                        /// </Explanation>

                        pRow["End1OID"] = pBusBarLV.OID.ToString();
                        pRow["End1Type"] = pBusBarLV.Class.ObjectClassID.ToString();
                        pRow["End1No"] = pBusBarLV.OID.ToString();
                        pRow["End1Position"] = "1";
                        pRow["End1BusBarLV"] = pBusBarLV.OID.ToString();
                        IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                        IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                        IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                        IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                        BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 1, pBusBarLV);
                        getPillarFeature(pBusBarLV, ref pRow, 1);
                        bGetConnectedFeature = true;
                        #endregion
                    }
                }
                else
                {
                    #region if connected feature is SP
                    IFeature pSPFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pPoint, 0.014,
                                                                                pServicePointFeatureLayer, "");
                    if (pSPFeature != null)
                    {
                        //Get ServicePoint Details
                        string strCOID = pSPFeature.get_Value(iCOID).ToString();
                        pRow["End1Type"] = pSPFeature.Class.ObjectClassID.ToString();
                        pRow["End1OID"] = pSPFeature.OID.ToString();
                        pRow["End1No"] = "";
                        pRow["End1Position"] = "";
                        pRow["End1FacilityContainsBusCoupler"] = "N";
                        //pRow["UnitConsumption"] = get_LOAD(strCOID);
                        pRow["UnitConsumption"] = "10";
                        pRow["End1FacilityOID"] = "SP";
                        pRow["End1FacilityName"] = "SP";
                        pRow["End1FacilityType"] = "SP";
                        //   pRow["End2FacilityTypeOf"] = "SP";
                        pRow["COID"] = strCOID;
                        bGetConnectedFeature = true;
                    }
                    else
                    {
                        //error table entry for no BusBarLV
                        #region
                        //iErrorCount++;
                        //DataRow pErrorRow = dtErrorResults.NewRow();
                        //pErrorRow["SrNo"] = iErrorCount.ToString();
                        //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        //pErrorRow["LayerName"] = "LTCable";
                        //pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        //pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        //dtErrorResults.Rows.Add(pErrorRow);
                        //_bSnappingErrorCheck = true;
                        #endregion
                    }
                    #endregion

                    #region Commented 18 march 2010
                    ////IFeature pNetJunctFeature = getIntersectingFeature(_pPoint, pNetJunctionFeatureLayer, "");
                    //IFeature pNetJunctFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pPoint, 0.02,
                    //                                                                    pNetJunctionFeatureLayer, "");
                    //if (pNetJunctFeature != null)
                    //{
                    //    string sLTCableOID = pCurrentLTCableFeature.OID.ToString();
                    //    IFeature pLTCable = getFeature4GenericCollection(pNetJunctFeature, pLTCableFeatureLayer, sLTCableOID);
                    //    if (pLTCable != null)
                    //    {
                    //        //IFeature pSPFeature = getIntersectingFeature(pLTCable, pServicePointFeatureLayer, "");
                    //        IFeature pSPFeature = Setting_Layers_Path.GetFirstBufferedFeature(pLTCable.Shape, 0.02,
                    //                                                                pServicePointFeatureLayer, "");
                    //        if (pSPFeature != null)
                    //        {
                    //            //get COID Details
                    //            string sCOID = pSPFeature.get_Value(iCOID).ToString();
                    //            //Add new row in TappedLTCable DataTable
                    //            DataRow pTappedCableRow = dtTappedLTCable.NewRow();
                    //            pTappedCableRow["SrNo"] = "";
                    //            pTappedCableRow["End1OID"] = pLTCable.OID.ToString();
                    //            pTappedCableRow["CurrentFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                    //            pTappedCableRow["COID"] = sCOID;
                    //            //pRow["UnitConsumption"] = get_LOAD(COID);
                    //            pTappedCableRow["UnitConsumption"] = "10";
                    //            pTappedCableRow.AcceptChanges();
                    //            dtTappedLTCable.Rows.Add(pTappedCableRow);
                    //            //Return False for TracedResults Table
                    //            bGetConnectedFeature = false;
                    //        }
                    //        else
                    //        {
                    //            iErrorCount++;
                    //            DataRow pErrorRow = dtErrorResults.NewRow();
                    //            pErrorRow["SrNo"] = iErrorCount.ToString();
                    //            pErrorRow["LayerName"] = "LT / SP ";
                    //            pErrorRow["ErrorType"] = "No SP found connected to LTcable";
                    //            pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                    //            dtErrorResults.Rows.Add(pErrorRow);
                    //        }
                    //    }
                    //}
                    //else
                    //{
                    //    #region if connected feature is SP
                    //    IFeature pSPFeature = Setting_Layers_Path.GetFirstBufferedFeature(pCurrentLTCableFeature.Shape, 0.015,
                    //                                                                pServicePointFeatureLayer, "");
                    //    if (pSPFeature != null)
                    //    {
                    //        //Get ServicePoint Details
                    //        string COID = pSPFeature.get_Value(iCOID).ToString();
                    //        pRow["End1Type"] = pSPFeature.Class.ObjectClassID.ToString();
                    //        pRow["End1OID"] = pSPFeature.OID.ToString();
                    //        pRow["End1No"] = "";
                    //        pRow["End1Position"] = "";
                    //        pRow["End1FacilityContainsBusCoupler"] = "N";
                    //        //   pRow["UnitConsumption"] = get_LOAD(COID);
                    //        pRow["UnitConsumption"] = "10";
                    //        pRow["End1FacilityOID"] = "SP";
                    //        pRow["End1FacilityName"] = "SP";
                    //        pRow["End1FacilityType"] = "SP";
                    //        //   pRow["End2FacilityTypeOf"] = "SP";
                    //        pRow["COID"] = COID;
                    //        bGetConnectedFeature = true;
                    //        //If the LTCable is Tapped 

                    //    }
                    //    #endregion
                    //    else
                    //    {
                    //        iErrorCount++;
                    //        DataRow pErrorRow = dtErrorResults.NewRow();
                    //        pErrorRow["SrNo"] = iErrorCount.ToString();
                    //        pErrorRow["LayerName"] = "LT / * ";
                    //        pErrorRow["ErrorType"] = "No End Feature found connected to LTcable";
                    //        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                    //        dtErrorResults.Rows.Add(pErrorRow);
                    //    }
                    //    #region If feature is connected to TerminationPoint
                    //    //IFeature pTerminationPoint = getIntersectingFeature(_pPoint, pTerminationPointFeatureLayer, "");
                    //    //if (pTerminationPoint != null)
                    //    //{
                    //    //    pRow["End1OID"] = pTerminationPoint.OID.ToString();
                    //    //    pRow["End1Type"] = pTerminationPoint.FeatureType.ToString();
                    //    //    pRow["End1No"] = pTerminationPoint.OID.ToString();
                    //    //    pRow["End1Position"] = "1";
                    //    //    IFeature pPillarFeature = getIntersectingFeature(pTerminationPoint, pPillarFeatureLayer, "");
                    //    //    if (pPillarFeature != null)
                    //    //    {
                    //    //        pRow["End1FacilityOID"] = pPillarFeature.OID.ToString();
                    //    //        pRow["End1FacilityType"] = "MP";
                    //    //        pRow["End1FacilityName"] = pPillarFeature.get_Value(iPillarLocation);
                    //    //    }
                    //    //    //If the connected feature is a busdropper
                    //    //    pBDFeature = null;
                    //    //    pBDFeature = getIntersectingFeature(pTerminationPoint, pBusBarFeatureLayer, "SUBTYPECD = 4");
                    //    //    if (pBDFeature != null)
                    //    //    {
                    //    //        pBusBarLV = null;
                    //    //        pBusBarLV = getIntersectingFeature(pBDFeature, pBusBarFeatureLayer, "SUBTYPECD = 3");
                    //    //        if (pBusBarLV != null)
                    //    //        {
                    //    //            pRow["End1BusBarLV"] = pBusBarLV.OID.ToString();
                    //    //            pRow["End1FacilityContainsBusCoupler"] = "N";
                    //    //            pRow["End1BC1OID"] = "0";
                    //    //            pRow["End1BC1Position"] = "0";
                    //    //            pRow["End1BC2OID"] = "0";
                    //    //            pRow["End1BC2Position"] = "0";
                    //    //            bGetConnectedFeature = true;
                    //    //        }
                    //    //        else
                    //    //        {
                    //    //            //error table entry for no BusBarLV
                    //    //            iErrorCount++;
                    //    //            DataRow pErrorRow = dtErrorResults.NewRow();
                    //    //            pErrorRow["SrNo"] = iErrorCount.ToString();
                    //    //            pErrorRow["LayerName"] = "BusDropper / BusBarLV";
                    //    //            pErrorRow["ErrorType"] = "No BusBarLV found connected to BusDropper";
                    //    //            pErrorRow["OID"] = pBDFeature.OID.ToString();
                    //    //            dtErrorResults.Rows.Add(pErrorRow);
                    //    //        }
                    //    //    }
                    //    //    else
                    //    //    {
                    //    //        //error table entry for no BDFeature 
                    //    //        iErrorCount++;
                    //    //        DataRow pErrorRow = dtErrorResults.NewRow();
                    //    //        pErrorRow["SrNo"] = iErrorCount.ToString();
                    //    //        pErrorRow["LayerName"] = "Termination / BusDropper";
                    //    //        pErrorRow["ErrorType"] = "No BusDropper found connected to Termination Point";
                    //    //        pErrorRow["OID"] = pTerminationPoint.OID.ToString();
                    //    //        dtErrorResults.Rows.Add(pErrorRow);
                    //    //    }
                    //    //}
                    //    #endregion
                    //}\
                    #endregion
                }
                #endregion Find connected BusBar-LV
            }            
            return bGetConnectedFeature;
        }

        private bool getEnd2ConnectedFeature(IPoint _pPoint, IFeature pCurrentLTCableFeature, ref DataRow pRow, ref bool _bSnappingErrorCheck)
         {
            bool bGetConnectedFeature = false;
            //If connected Feature is BusDropper
            //IFeature pBDFeature = getIntersectingFeature(_pPoint, pBusBarFeatureLayer, "SUBTYPECD = 4");
            if (pCurrentLTCableFeature.OID.ToString() == "2300955")
            {
            }
            IFeature pBDFeature = Setting_Layers_Path.GetFirstBufferedFeature((IGeometry)_pPoint, 0.00101,
                                                                             pBusBarFeatureLayer, "SUBTYPECD = 4");
            if (pBDFeature != null)
            {
                #region  Chking & getting Fuse details
                IFeature pFuse = getIntersectingFeature(pBDFeature, pFuseFeatureLayer, "");
                if (pFuse != null)
                {
                    pRow["End2OID"] = pFuse.OID.ToString();
                    pRow["End2Type"] = pFuse.Class.ObjectClassID.ToString();
                    pRow["End2No"] = pFuse.OID.ToString();
                    pRow["End2Position"] = pFuse.get_Value(iFusePositionFldInd).ToString();
                    getPillarFeature(pFuse, ref pRow, 2); 
                    pBDFeature = getBusDropperFeature(pFuse, pBDFeature.OID.ToString());
                    #region Chking & getting BD details
                    if (pBDFeature != null)
                    {
                        //IFeature pBusBarLV = getIntersectingFeature(pBDFeature,pBusBarFeatureLayer,
                        //"SUBTYPECD = 3");
                        IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.008,
                                                                            pBusBarFeatureLayer,"SUBTYPECD = 3");
                        #region Chking & getting BC details
                        if (pBusBarLV != null)
                        {
                            pRow["End2BusBarLV"] = pBusBarLV.OID.ToString();
                            IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                            IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                            IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                            IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                            BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 2, pBusBarLV);
                           // pRow["COID"] = "0";
                           // pRow["UnitConsumption"] = "0";
                            bGetConnectedFeature = true;
                            //fill Fuse property columns with BusDropper properties 
                            //pRow["End2OID"] = pBDFeature.OID.ToString();
                            //pRow["End2Type"] = pBDFeature.Class.ObjectClassID.ToString();
                            //pRow["End2No"] = pBDFeature.OID.ToString();
                            //pRow["End2Position"] = "1"; //Assuming that BusDropper can never be OFF
                        }
                        else
                        {
                            #region Snapping Error
                            iErrorCount++;
                            DataRow pErrorRow = dtErrorResults.NewRow();
                            pErrorRow["SrNo"] = iErrorCount.ToString();
                            pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                            pErrorRow["LayerName"] = "BusDropper";
                            pErrorRow["ErrorType"] = "No Feature found connected to BusDropper";
                            pErrorRow["OID"] = pBDFeature.OID.ToString();
                            dtErrorResults.Rows.Add(pErrorRow);
                            _bSnappingErrorCheck = true;
                            #endregion
                        }
                        #endregion  Chking & getting BC details
                    }
                    else
                    {
                        #region Snapping Error
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "Fuse";
                        pErrorRow["ErrorType"] = "No Feature found connected to Fuse";
                        pErrorRow["OID"] = pFuse.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        _bSnappingErrorCheck = true;
                        #endregion
                    }
                    #endregion  Chking & getting BD details
                }
                else
                {
                    //IFeature pBusBarLV = getIntersectingFeature(pBDFeature,pBusBarFeatureLayer,
                                                                            //"SUBTYPECD = 3");
                    IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.014,
                                                                            pBusBarFeatureLayer, "SUBTYPECD = 3");
                    #region Chking & getting BC details
                    if (pBusBarLV != null)
                    {
                        pRow["End2OID"] = pBDFeature.OID.ToString();
                        pRow["End2Type"] = pBDFeature.Class.ObjectClassID.ToString();
                        pRow["End2No"] = pBDFeature.OID.ToString();
                        pRow["End2Position"] = "1";
                        getPillarFeature(pBusBarLV, ref pRow, 2);

                        pRow["End2BusBarLV"] = pBusBarLV.OID.ToString();
                        IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                        IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                        IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                        IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);

                        BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 2, pBusBarLV);
                     //   pRow["COID"] = "0";
                     //   pRow["UnitConsumption"] = "0";
                        bGetConnectedFeature = true;
                    }
                    else
                    {
                        #region Snapping Error
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "LTCable";
                        pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        _bSnappingErrorCheck = true;
                        #endregion
                    }
                    #endregion  Chking & getting BC details                    
                }
                #endregion  Chking & getting Fuse details
            }
            else
            { 
                #region if connected feature is SP
                 IFeature pSPFeature = getIntersectingFeature(_pPoint, pServicePointFeatureLayer, "");
                 if (pSPFeature != null)
                 {
                     //Get ServicePoint Details
                     string strCOID = pSPFeature.get_Value(iCOID).ToString();
                     pRow["End2Type"] = pSPFeature.Class.ObjectClassID.ToString();
                     pRow["End2OID"] = pSPFeature.OID.ToString();
                     pRow["End2No"] = "";
                     pRow["End2Position"] = "";
                     pRow["End2FacilityContainsBusCoupler"] = "N";
                     //pRow["UnitConsumption"] = get_LOAD(strCOID);
                     pRow["UnitConsumption"] = "10";
                     pRow["End2FacilityOID"] = "SP";
                     pRow["End2FacilityName"] = "SP";
                     pRow["End2FacilityType"] = "SP";
                     //   pRow["End2FacilityTypeOf"] = "SP";
                     pRow["COID"] = strCOID;
                     bGetConnectedFeature = true;
                 }
                #endregion
                 else
                 {
                     #region Connected to BusBar-LV
                     //IFeature pBusBarLV = getIntersectingFeature(_pPoint, pBusBarFeatureLayer,
                                                                                 //"SUBTYPECD = 3");
                     IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(_pPoint, 0.008,
                                                                            pBusBarFeatureLayer, "SUBTYPECD = 3");
                     if (pBusBarLV != null)
                     {
                         #region Chk for FUMP
                         IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBusBarLV.Shape,
                                                                            pPillarFeatureLayer,
                                                                            esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                            "");
                         int iFUMPChk = 0;
                         if (pPillarFeature != null)
                         {

                             string strSubType = CommonGISFunc.GetFieldValueFN(pPillarFeature, "SUBTYPECD", true);
                             if (strSubType == "4")
                             {
                                 iFUMPChk = 1;
                             }
                         }
                         if (iFUMPChk == 1)
                         {
                             GetFUMPDetails(pBusBarLV, ref pRow, 2);
                         }
                         #endregion
                         else
                         {
                             pRow["End2OID"] = pBusBarLV.OID.ToString();
                             pRow["End2Type"] = pBusBarLV.Class.ObjectClassID.ToString();
                             pRow["End2No"] = pBusBarLV.OID.ToString();
                             pRow["End2Position"] = "1";
                             pRow["End2BusBarLV"] = pBusBarLV.OID.ToString();
                             IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                             IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                             IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                             IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                             BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, 2, pBusBarLV);
                             getPillarFeature(pBusBarLV, ref pRow, 2);
                             bGetConnectedFeature = true;
                         }
                     }
                     else
                     {
                         //error table entry for no BusBarLV 
                         #region Snapping Error commented
                         //  iErrorCount++;
                         //DataRow pErrorRow = dtErrorResults.NewRow();
                         //pErrorRow["SrNo"] = iErrorCount.ToString();
                         //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                         //pErrorRow["LayerName"] = "LTCable";
                         //pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                         //pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                         //dtErrorResults.Rows.Add(pErrorRow);
                         //_bSnappingErrorCheck = true;                                                                                                                      
                         #endregion
                     }
                     #endregion
                 }                                
            }
            return bGetConnectedFeature;
        }

        private void GetFUMPDetails(IFeature _pBusBarLV, ref DataRow _pRow, int _iEnd)
        {
            try
            {
                IPolyline pBBLVPolyLine = (IPolyline)_pBusBarLV.Shape;
                IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);

                FUMPdetails(pFirstPoint, pLastPoint, ref _pRow, _iEnd, _pBusBarLV);
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
        }

        private void FUMPdetails(IPoint pFirstPoint, IPoint pLastPoint, ref DataRow _pRow, int _iEnd, IFeature pBBFeature)
        {
            try
            {                
                int iflag_fuse = 0;
                int iFlag = 0;

                GenericCollection pFuseFeatureGenericCollection = new GenericCollection();
                CommonGISFunc.AllSpatialQueryFeatures(pBBFeature.Shape, pFuseFeatureLayer,
                                                        esriSpatialRelEnum.esriSpatialRelIntersects, "",
                                                        ref pFuseFeatureGenericCollection);                
                if (pFuseFeatureGenericCollection.Count == 2)
                {
                    setUniqueBBOID4TwoFuses(pFuseFeatureGenericCollection, pBBFeature);                    
                }              
                else
                {
                    foreach (IFeature pFusefeat in pFuseFeatureGenericCollection)
                    {
                        if (pFusefeat.get_Value(iFusePositionFldInd).ToString() == "1")
                        {
                            IFeature pBusBarFeature = getFeature4GenericCollection(pFusefeat, pBusBarFeatureLayer, pBBFeature.OID.ToString());
                            if (pBusBarFeature != null)
                            {
                                IFeature pFuse2feature = getFeature4GenericCollection(pBusBarFeature, pFuseFeatureLayer, pFusefeat.OID.ToString());
                                if (pFuse2feature.get_Value(iFusePositionFldInd).ToString() == "1")
                                {
                                    /////send midle BB
                                    GenericCollection pFuseCollection = new GenericCollection();
                                    CommonGISFunc.AllSpatialQueryFeatures(pBusBarFeature.Shape, pFuseFeatureLayer,
                                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "",
                                                                    ref pFuseCollection);
                                    if (pFuseCollection.Count == 2)
                                    {
                                        setUniqueBBOID4TwoFuses(pFuseCollection, pBusBarFeature);
                                    }
                                }
                                else
                                {
                                    //For the First Feature
                                    #region For _pFirstPoint Feature
                                    iFlag = 1; iflag_fuse = 1;
                                    getDetails4pointFeature(pFirstPoint, ref _pRow, pBBFeature, _iEnd, iFlag, iflag_fuse);
                                    #endregion

                                    #region For pSecondPoint Feature
                                    iFlag = 0; iflag_fuse = 0;
                                    getDetails4pointFeature(pLastPoint, ref _pRow, pBBFeature, _iEnd, iFlag, iflag_fuse);
                                    #endregion
                                }
                            }
                        }
                        else
                        {
                            //For the First Feature
                            #region For _pFirstPoint Feature
                            iFlag = 1; iflag_fuse = 1;
                            getDetails4pointFeature(pFirstPoint, ref _pRow, pBBFeature, _iEnd, iFlag, iflag_fuse);
                            #endregion

                            #region For pSecondPoint Feature
                            iFlag = 0; iflag_fuse = 0;
                            getDetails4pointFeature(pLastPoint, ref _pRow, pBBFeature, _iEnd, iFlag, iflag_fuse);
                            #endregion
                        }

                    }
                }

            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
        }

        private void setUniqueBBOID4TwoFuses(GenericCollection pFuseFeatureGenericCollection, IFeature pBBFeature)
        {
           string sFusePosition = "";
           try
           {
               #region
               foreach (IFeature pFuseFeat in pFuseFeatureGenericCollection)
               {
                   sFusePosition += pFuseFeat.get_Value(iFusePositionFldInd).ToString();
               }
               bool bFusepositionflag = true;
               if (sFusePosition.Length > 0)
                   bFusepositionflag = sFusePosition.Contains("0");

               if (pFuseFeatureGenericCollection.Count == 2 & bFusepositionflag == false)
               {
                   ArrayList arrBB_OIDS = new ArrayList();
                   arrBB_OIDS.Add(pBBFeature.OID.ToString());
                   #region For Two FuseFeatures in FUMP
                   foreach (IFeature pFuseFeat in pFuseFeatureGenericCollection)
                   {
                       GenericCollection pTempBBFeatCollection = new GenericCollection();
                       CommonGISFunc.AllSpatialQueryFeatures(pFuseFeat.Shape, pBusBarFeatureLayer,
                                                                   esriSpatialRelEnum.esriSpatialRelIntersects, "",
                                                                   ref pTempBBFeatCollection);

                       foreach (IFeature pTempBBFeat in pTempBBFeatCollection)
                       {
                           if (pTempBBFeat != null)
                           {
                               if (arrBB_OIDS.Contains(pTempBBFeat.OID.ToString()))
                               {
                               }
                               else
                                   arrBB_OIDS.Add(pTempBBFeat.OID.ToString());
                           }
                       }
                   }
                   if (arrBB_OIDS.Count == 3)
                   {
                       IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBBFeature.Shape, pPillarFeatureLayer,
                                                                                        esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                                           "");
                       if (!arrFUMPTypUniPillarID.Contains(pPillarFeature.OID.ToString()))
                       {
                           arrFUMPTypUniPillarID.Add(pPillarFeature.OID.ToString());
                           DataRow pRow = dtFUMPPillarTable.NewRow();
                           pRow["PillarOID"] = pPillarFeature.OID.ToString();
                           pRow["BusBarLV-I"] = arrBB_OIDS[0].ToString();
                           pRow["BusBarLV-II"] = arrBB_OIDS[1].ToString();
                           pRow["BusBarLV-III"] = arrBB_OIDS[2].ToString();
                           dtFUMPPillarTable.Rows.Add(pRow);
                       }
                   }


                   #endregion
               }
               #endregion
           }
           catch (Exception err)
           {
               MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
           }
           finally
           {
               #region
               #endregion
           }
        }

        private void getDetails4pointFeature(IPoint _pPoint, ref DataRow _pRow, IFeature pBBFeature, int _iEnd, int iFlag, int iflag_fuse )
        {
            string sModifiedBBLVOID = string.Empty;
            try
            {
                IFeature pFuseFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)_pPoint,pFuseFeatureLayer,
                                                                esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                "");
                if (pFuseFeature != null)
                {                   
                    _pRow["End" + _iEnd.ToString() + "BC1OID"] = pFuseFeature.OID.ToString();
                    _pRow["End" + _iEnd.ToString() + "BC1Position"] = pFuseFeature.get_Value(iFusePositionFldInd).ToString();

                    if (pFuseFeature.get_Value(iFusePositionFldInd).ToString() == "1")
                    {
                        // setting BBOID
                        sModifiedBBLVOID = setFUMPBusBarLVOID(pFuseFeature, pBBFeature, iflag_fuse);
                        _pRow["End" + _iEnd.ToString() + "BusBarLV"] = sModifiedBBLVOID;
                        _pRow["Remarks" + _iEnd.ToString()] = String.Empty;
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
        }

        private string setFUMPBusBarLVOID(IFeature pFuseFeature, IFeature pBBFeature, int _iFlag_Fuse)
        {
            string sReturn = null;
            try
            {
                string sBBOID = pBBFeature.OID.ToString();
                IFeature p2BBLV = getFeature4GenericCollection(pFuseFeature, pBusBarFeatureLayer, sBBOID);
                if (p2BBLV != null)
                {
                    string str2BBLV = p2BBLV.OID.ToString();

                    if (_iFlag_Fuse == 0)
                        sReturn = str2BBLV + "_" + sBBOID;
                    else
                        sReturn = sBBOID + "_" + str2BBLV;
                } 
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
            return sReturn;
        }        
             
        private string get_LOAD(string sCOID)
        {
            DataSet ds = null;
            string _connectionString = "Data Source=gisdb;User Id= vedgo;Password=41003756";
            using (OracleConnection _oraconn = new OracleConnection(_connectionString))
            {
                try
                {
                    _oraconn.Open();
                    string strSQL = "select * from RELGDB.COID_CONSUMPTION_VIEW where COID = '" + sCOID + "'";
                    using (OracleCommand cmd = new OracleCommand(strSQL, _oraconn))
                    {
                        cmd.CommandType = CommandType.Text;
                        using (OracleDataAdapter da = new OracleDataAdapter(cmd))
                        {
                            ds = new DataSet("COID_CONSUMPTION");
                            da.Fill(ds);
                        }
                    }
                    _oraconn.Close();
                }
                catch (OracleException ex)
                {
                    MessageBox.Show("Database error: " + ex.Message.ToString());
                }
                catch (Exception err)
                {
                    MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + 
                                                                err.StackTrace + "\nTargetSite: " + err.TargetSite);
                }
            }            
            string sLOAD = "";
            try
            {                            
                int cnt = ds.Tables[0].Rows.Count;
                if (cnt > 0)
                {
                    foreach (DataRow pRow in ds.Tables[0].Rows)
                    {
                        sLOAD = pRow["CONNECTED_LOAD"].ToString();
                    }                    
                }
                else
                    sLOAD= "0";
            }            
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + 
                                                                err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            return sLOAD;
        }                   

        private IFeature getBusDropperFeature(IFeature _pFeature, string _sBDOID)
        {
	        IFeature pBusDropperFeature = null;	
	        GenericCollection pBusDroperFeatureCol = new GenericCollection();
	        CommonGISFunc.AllSpatialQueryFeatures(_pFeature.Shape, pBusBarFeatureLayer, 
											        esriSpatialRelEnum.esriSpatialRelIntersects, "", 
											        ref pBusDroperFeatureCol);
	        //if BusDropperCollection contains more than one feature
            if (pBusDroperFeatureCol.Count > 1)
	        {
                foreach (IFeature pBDFeature in pBusDroperFeatureCol)
		        {
			        if(pBDFeature.OID.ToString() != _sBDOID)
			        {
				        pBusDropperFeature = pBDFeature;
			        }
		        }
	        }
	        //return BusDropper Feature
	        return pBusDropperFeature;	
        }      

        private IFeature getIntersectingFeature(IPoint _pPoint, IFeatureLayer _pIntersectionFeatureLayer, string _sExp)
        {
            IFeature pFeature = null;
            pFeature = CommonGISFunc.FirstSpatialQueryFeature((IPoint)_pPoint, _pIntersectionFeatureLayer,
                                                            esriSpatialRelEnum.esriSpatialRelIntersects, _sExp);
            return pFeature;
        }

        private IFeature getIntersectingFeature(IFeature _pFeature, IFeatureLayer _pIntersectionFeatureLayer, string _sExp)
        {
            IFeature pFeature = null;
            pFeature = CommonGISFunc.FirstSpatialQueryFeature(_pFeature.Shape, _pIntersectionFeatureLayer,
                                                            esriSpatialRelEnum.esriSpatialRelIntersects, _sExp);
            return pFeature;
        }
        
        private void getPillarFeature(IFeature pFeature, ref DataRow pRow, int iEnd)
        {
            IFeature pPillarFeature = null;
            try
            {
                pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pFeature.Shape,
                                                                                pPillarFeatureLayer,
                                                                                esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                                "");
                if (pPillarFeature != null)
                {
                    pRow["End" + iEnd.ToString() + "FacilityOID"] = pPillarFeature.OID.ToString();
                    string strSubType = CommonGISFunc.GetFieldValueFN(pPillarFeature, "SUBTYPECD", true);
                    if (strSubType == "3")
                        pRow["End" + iEnd.ToString() + "FacilityType"] = "MiniPillar";
                    if (strSubType == "2")
                    {
                        string strTotalNoofCkt = pPillarFeature.get_Value(iTotalCktNos).ToString();
                        //pRow["End" + iEnd.ToString() + "FacilityType"] = "LTPillar";

                        pRow["End" + iEnd.ToString() + "FacilityType"] = pPillarFeature.get_Value(iTotalCktNos).ToString() + " W";
                    }
                    if (strSubType == "4")
                        pRow["End" + iEnd.ToString() + "FacilityType"] = "FUMP";
                    
                    pRow["End" + iEnd.ToString() + "FacilityName"] = pPillarFeature.get_Value(iPillarLocation).ToString();
                }
            }            
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite +
                                                 "\nTarget Pillar Feature: " + pPillarFeature.OID.ToString());
            }
            finally
            {
                #region
                #endregion
            }
        }

        private void getPillarDetails(IFeature pFeature)
        {
            IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pFeature.Shape,
                                                                            pPillarFeatureLayer,
                                                                            esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                            "");
            try
            {
                //if (pPillarFeature != null)
                //{
                //    if(!aryPillars.Contains(pPillarFeature.OID.ToString()))
                //    {
                //        aryPillars.Add(pPillarFeature.OID.ToString());
                //        DataRow pRow = dtPillarTable.NewRow();
                //        pRow["PillarOID"] = pPillarFeature.OID.ToString();
                //        pRow["BusBarLV-I"] = "";
                //    }
                //    else
                //    {
                //        string strPillarOID = pPillarFeature.OID.ToString();
                //        DataRow[] dRow = Search4PillarRow(dtPillarTable, strPillarOID);
                //        if (dRow.Length > 0)
                //        {
                //            foreach (DataRow _dRow in dRow)
                //            { 
                //                string strFeatOID = pFeature.OID.ToString();
                //                if (!_dRow["BusBarLV-I"].ToString() == strFeatOID)
                //                {
                //                    if (!_dRow["BusBarLV-II"].ToString() == strFeatOID)
                //                    {
                //                        if (!_dRow["BusBarLV-III"].ToString() == strFeatOID)
                //                        {
                //                            DataRow pRow = dtPillarTable.NewRow();
                //                            pRow["BusBarLV-I"].ToString() = strFeatOID;
                //                        }
                //                    }
                //                }                                
                //            }
                //        }
                        
                //        pRow["PillarOID"] = pPillarFeature.OID.ToString();
                //        pRow["BusBarLV-I"] = "";
                //    }
                //}
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region

                #endregion
            }
        }

        private DataRow[] Search4PillarRow(System.Data.DataTable LookInTable, string strOID)
        {
            DataRow[] FoundRows = null;
            try
            {
                //string sStrExpr = "(End1FacilityOID = '" + strID +"')" ;
                string sStrExpr = "(PillarOID = '" + strOID + "')";
                FoundRows = LookInTable.Select(sStrExpr);
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
            }
            return FoundRows;
        }

        private void BusCouplerdetails(IPoint _pFirst, IPoint _pLast, ref DataRow pRow1, int iEnd, IFeature pBBFeature)
        {
            IFeature pBCFeature = null;
            string sModifiedBBLVOID = null;
            int iflag_BB = 0;
            int iFlag = 0;

            GenericCollection pBC_FeatureGenericCollection = new GenericCollection();
            CommonGISFunc.AllSpatialQueryFeatures(pBBFeature.Shape, pSwitchFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "SUBTYPECD = 5",
                                                    ref pBC_FeatureGenericCollection);

            string sBCPosition = "";
            if (pBC_FeatureGenericCollection.Count == 2)
            { 
            }
            foreach (IFeature pBCFeat in pBC_FeatureGenericCollection)
            {
                sBCPosition += pBCFeat.get_Value(iSwitchLinkPositionFldInd).ToString();
            }

            bool sBCpositionflag = true;
            if (sBCPosition.Length > 0)
                sBCpositionflag = sBCPosition.Contains("0");
            if (pBC_FeatureGenericCollection.Count == 2 && sBCpositionflag == false)
            {
                #region For 2 BC
                ArrayList arrBB_OIDS = new ArrayList();
                arrBB_OIDS.Add(pBBFeature.OID.ToString());
                //IFeature pTempBBFeat = null;
                foreach (IFeature pBCFeat in pBC_FeatureGenericCollection)
                {
                    //pTempBBFeat = CommonGISFunc.FirstSpatialQueryFeature(pBCFeat.Shape, pBusBarFeatureLayer,
                    GenericCollection pTempBBFeatCollection = new GenericCollection();
                    CommonGISFunc.AllSpatialQueryFeatures(pBCFeat.Shape, pBusBarFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "",
                                                    ref pTempBBFeatCollection);
                    foreach (IFeature pTempBBFeat in pTempBBFeatCollection)
                    {
                        if (pTempBBFeat != null)
                        {
                            if (arrBB_OIDS.Contains(pTempBBFeat.OID.ToString()))
                            {
                            }
                            else
                                arrBB_OIDS.Add(pTempBBFeat.OID.ToString());
                        }
                        //pTempBBFeat = null;
                    }
                }
                if (arrBB_OIDS.Count == 3)
                {

                    IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBBFeature.Shape, pPillarFeatureLayer,
                                                                                     esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                                                  "");
                    if (!arrUniPillarID.Contains(pPillarFeature.OID.ToString()))
                    {
                        arrUniPillarID.Add(pPillarFeature.OID.ToString());
                        DataRow pRow = dtPillarTable.NewRow();
                        pRow["PillarOID"] = pPillarFeature.OID.ToString();
                        pRow["BusBarLV-I"] = arrBB_OIDS[0].ToString();
                        pRow["BusBarLV-II"] = arrBB_OIDS[1].ToString();
                        pRow["BusBarLV-III"] = arrBB_OIDS[2].ToString();
                        dtPillarTable.Rows.Add(pRow);
                    }
                }
                #endregion
            }
            else
            {
                #region For _pFirst Point Feature
                pBCFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)_pFirst,
                                                                                pSwitchFeatureLayer,
                                                                                esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                                "SUBTYPECD = 5");
                if (pBCFeature != null)
                {
                    iflag_BB = 1;
                    iFlag = 1;
                    // pRow["End2FacilityContainsBusCoupler"] = "Y";
                    pRow1["End" + iEnd.ToString() + "BC1OID"] = pBCFeature.OID.ToString();
                    pRow1["End" + iEnd.ToString() + "BC1Position"] = pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString();
                    if (pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString() == "1")
                    {
                        // setting BBOID
                        sModifiedBBLVOID = setBusBarLVOID(pBCFeature, pBBFeature, iflag_BB);
                        pRow1["End" + iEnd.ToString() + "BusBarLV"] = sModifiedBBLVOID;
                        pRow1["Remarks" + iEnd.ToString()] = String.Empty;
                    }
                    else
                        pRow1["Remarks" + iEnd.ToString()] = "Partial Pillar";
                }
                else
                {
                    //  pRow["End2FacilityContainsBusCoupler"] = "N";
                    pRow1["End" + iEnd.ToString() + "BC1OID"] = "";
                    pRow1["End" + iEnd.ToString() + "BC1Position"] = "";
                }
                #endregion

                pBCFeature = null;

                #region For _pLast Point Feature
                pBCFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)_pLast,
                                                                               pSwitchFeatureLayer,
                                                                               esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                               "SUBTYPECD = 5");
                if (pBCFeature != null)
                {
                    iFlag = 1;
                    // pRow["End2FacilityContainsBusCoupler"] = "Y";
                    pRow1["End" + iEnd.ToString() + "BC2OID"] = pBCFeature.OID.ToString();
                    pRow1["End" + iEnd.ToString() + "BC2Position"] = pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString();
                    if (pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString() == "1")
                    {
                        iflag_BB = 0;
                        sModifiedBBLVOID = setBusBarLVOID(pBCFeature, pBBFeature, iflag_BB);
                        pRow1["End" + iEnd.ToString() + "BusBarLV"] = sModifiedBBLVOID;
                        pRow1["Remarks" + iEnd.ToString()] = "";
                    }
                    else
                        pRow1["Remarks" + iEnd.ToString()] = "Partial Pillar";
                }
                else
                {
                    //  pRow["End2FacilityContainsBusCoupler"] = "N";
                    pRow1["End" + iEnd.ToString() + "BC2OID"] = "";
                    pRow1["End" + iEnd.ToString() + "BC2Position"] = "";
                }
                #endregion

                if (iFlag == 0)
                {
                    pRow1["End" + iEnd.ToString() + "FacilityContainsBusCoupler"] = "N";
                    //pRow1["Remarks"] = "Pillar does not contains BusCoupler";
                }
                else
                {
                    pRow1["End" + iEnd.ToString() + "FacilityContainsBusCoupler"] = "Y";
                    // pRow1["Remarks"] = "Partial Pillar";
                }
            }
        }

        //private void BusCouplerdetails(IPoint _pFirst, IPoint _pLast, ref DataRow pRow1, int iEnd, IFeature pBBFeature)
        //{
        //    IFeature pBCFeature = null;
        //    int iFlag = 0;
        //    try
        //    {
        //        //For First Point
        //        #region For First Point
        //        pBCFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)_pFirst,
        //                                                            pSwitchFeatureLayer,
        //                                                            esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                            "SUBTYPECD = 5");
        //        if (pBCFeature != null)
        //        {                    
        //            iFlag = 1;
        //            // pRow["End2FacilityContainsBusCoupler"] = "Y";
        //            pRow1["End" + iEnd.ToString() + "BC1OID"] = pBCFeature.OID.ToString();
        //            pRow1["End" + iEnd.ToString() + "BC1Position"] = pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString();
        //            if (pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString() == "1")
        //            {                                                                      
        //                pRow1["Remarks" + iEnd.ToString()] = "";
        //            }
        //            else
        //                pRow1["Remarks" + iEnd.ToString()] = "Partial Pillar";
        //        }
        //        else
        //        {
        //            //  pRow["End2FacilityContainsBusCoupler"] = "N";
        //            pRow1["End" + iEnd.ToString() + "BC1OID"] = "";
        //            pRow1["End" + iEnd.ToString() + "BC1Position"] = "";
        //        }
        //        #endregion

        //        pBCFeature = null;

        //        #region For Second Point
        //        pBCFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)_pLast,
        //                                                           pSwitchFeatureLayer,
        //                                                           esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                           "SUBTYPECD = 5");
        //        if (pBCFeature != null)
        //        {
        //            iFlag = 1;
        //            // pRow["End2FacilityContainsBusCoupler"] = "Y";
        //            pRow1["End" + iEnd.ToString() + "BC2OID"] = pBCFeature.OID.ToString();
        //            pRow1["End" + iEnd.ToString() + "BC2Position"] = pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString();
        //            if (pBCFeature.get_Value(iSwitchLinkPositionFldInd).ToString() == "1")
        //            {                        
        //                pRow1["Remarks" + iEnd.ToString()] = "";
        //            }
        //            else
        //                pRow1["Remarks" + iEnd.ToString()] = "Partial Pillar";
        //        }
        //        else
        //        {
        //            //  pRow["End2FacilityContainsBusCoupler"] = "N";
        //            pRow1["End" + iEnd.ToString() + "BC2OID"] = "";
        //            pRow1["End" + iEnd.ToString() + "BC2Position"] = "";
        //        }
        //        #endregion

        //        #region Update FacilityContainsBusCoupler Field by checking for BusCoupler Feature
        //        if (iFlag == 0)
        //        {
        //            pRow1["End" + iEnd.ToString() + "FacilityContainsBusCoupler"] = "N";                    
        //        }
        //        else
        //        {
        //            pRow1["End" + iEnd.ToString() + "FacilityContainsBusCoupler"] = "Y";                    
        //        }
        //        #endregion
        //    }
        //    catch (Exception err)
        //    {
        //        MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " +
        //                                                        err.StackTrace + "\nTargetSite: " + err.TargetSite);
        //    }
        //    finally
        //    {
        //    }
        //}

        private string setBusBarLVOID(IFeature pBCFeature, IFeature pBBFeature, int _iFlag_BB)
        {
            //string strLabelling = "";
            string sreturn = null;
            try
            {                 
                string sBBOID = pBBFeature.OID.ToString();
                IFeature p2BBLV = getFeature4GenericCollection(pBCFeature, pBusBarFeatureLayer, sBBOID);

                if (p2BBLV != null)
                {
                    string str2BBLV = p2BBLV.OID.ToString();

                    if (_iFlag_BB == 0)
                        sreturn = str2BBLV + "_" + sBBOID;
                    else
                        sreturn = sBBOID + "_" + str2BBLV;
                }                
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                
                #endregion
            }
            return sreturn;
        }

        private void setBusBarLVOID1223(IFeature pBCFeature, IFeature pBBFeature, ref DataRow pRow1, int _iFlag_BB)
        {
            //string strLabelling = "";
            //string sreturn = null;
            //try
            //{
                
            //    string sBBOID = pBBFeature.OID.ToString();
            //    IFeature p2BBLV = getFeature4GenericCollection(pBCFeature, pBusBarFeatureLayer, sBBOID);

            //    if (p2BBLV != null)
            //    {
            //        string str2BBLV = p2BBLV.OID.ToString();
            //        IFeature p2BC = getFeature4GenericCollection(p2BBLV, pSwitchFeatureLayer, str2BBLV);
            //        if (p2BC != null)
            //        {
            //            if (p2BC.get_Value(iSwitchLinkPositionFldInd).ToString() == "1")
            //            {
            //                // IFeature p3BusBarLV = 
            //            }
            //        }
            //        else
            //        {string sreturn = null;

            //            if (_iFlag_BB == 0)
            //                sreturn = p2BBLV.OID.ToString() + "_" + sBBOID;
            //            else
            //                sreturn = sBBOID + "_" + p2BBLV.OID.ToString();
            //        }
            //    }
            //    return sreturn;
            //}
            //catch (Exception err)
            //{
            //    MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
            //                                     "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            //}
            //finally
            //{
            //    #region

            //    #endregion
            //}
        }

        private IFeature getFeature4GenericCollection(IFeature _pFeature,IFeatureLayer _pFeatureLayer, string _sOID)
        {
            IFeature pIntersectingFeature = null;
            GenericCollection pFeatureGenericCollection = new GenericCollection();
            CommonGISFunc.AllSpatialQueryFeatures(_pFeature.Shape, _pFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, "",
                                                    ref pFeatureGenericCollection);
            //if BusDropperCollection contains more than one feature
            if (pFeatureGenericCollection.Count > 1)
            {
                foreach (IFeature pFeature in pFeatureGenericCollection)
                {
                    if (pFeature.OID.ToString() != _sOID)
                    {
                        pIntersectingFeature = pFeature;
                    }
                }
            }
            //return BusDropper Feature
            return pIntersectingFeature;
        }

        private IFeature getFeature4GenericCollection(IFeature _pFeature, IFeatureLayer _pFeatureLayer, 
                                                                            string _sOID, string _pQuery)
        {
            IFeature pIntersectingFeature = null;
            GenericCollection pFeatureGenericCollection = new GenericCollection();
            CommonGISFunc.AllSpatialQueryFeatures(_pFeature.Shape, _pFeatureLayer,
                                                    esriSpatialRelEnum.esriSpatialRelIntersects, _pQuery,
                                                    ref pFeatureGenericCollection);
            //if BusDropperCollection contains more than one feature
            if (pFeatureGenericCollection.Count > 1)
            {
                foreach (IFeature pFeature in pFeatureGenericCollection)
                {
                    if (pFeature.OID.ToString() != _sOID)
                    {
                        pIntersectingFeature = pFeature;
                    }
                }
            }
            //return BusDropper Feature
            return pIntersectingFeature;
        }

        private bool FindForDuplicateCable(System.Data.DataTable oLTCable_DataTable, DataRow pNewRow)
        {
            try
            {
                string sStrExpr = "(End1Type = '" + pNewRow["End1Type"].ToString() + "' AND End1OID = '" + pNewRow["End1OID"].ToString() + "' AND End2Type = '" + pNewRow["End2Type"].ToString() + "' AND End2OID = '" + pNewRow["End2OID"].ToString() + "') OR (End1Type = '" + pNewRow["End2Type"].ToString() + "' AND End1OID = '" + pNewRow["End2OID"].ToString() + "' AND End2Type = '" + pNewRow["End1Type"].ToString() + "' AND End2OID = '" + pNewRow["End1OID"].ToString() + "')";
                DataRow[] dRows = oLTCable_DataTable.Select(sStrExpr);
                if (dRows.Length > 0)
                {
                    IFeature pLTCableFeat = (IFeature)pNewRow["LTCableFeature"];
                    IFeature pLTCableFeat1 = (IFeature)dRows[0]["LTCableFeature"];
                    if (!(pLTCableFeat.OID.ToString().Equals(pLTCableFeat1.OID.ToString())))
                    {
                        sErrorMessage += "LT Cable with OID '" + pLTCableFeat.OID + "' is a duplicate piece" + Environment.NewLine;
                    }
                    return true;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " +
                                                                err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            return false;
        }              

        private bool GetTappedCableDetails(IPoint _pFirstPoint, IPoint _pLastPoint, IFeature pCurrentLTCableFeature, ref DataRow pRow)
        {
            bool bFirstTappedConnectedFeature = false;
            bool bSecondTappedConnectedFeature = false;
            bool bTappedConnectedFeature = false;
            IFeature pSPFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pFirstPoint, 0.014,
                                                                                pServicePointFeatureLayer, "");
            if (pSPFeature != null)
            {
                string strCOID = pSPFeature.get_Value(iCOID).ToString();
                pRow["COID"] = strCOID;
                //pRow["UnitConsumption"] = get_LOAD(strCOID);
                pRow["UnitConsumption"] = "10";
                bFirstTappedConnectedFeature = true;
            }
            else
            {
                IFeature pNetJunctFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pFirstPoint, 0.014,
                                                                                pNetJunctionFeatureLayer, "");
                if (pNetJunctFeature != null)
                {
                    string sLTCableOID = pCurrentLTCableFeature.OID.ToString();
                    IFeature pLTCable = getFeature4GenericCollection(pNetJunctFeature, pLTCableFeatureLayer,
                                                                                                    sLTCableOID);
                    if (pLTCable != null)
                    {
                        pRow["SrNo"] = "";
                        pRow["End1OID"] = pLTCable.OID.ToString();
                        pRow["CurrentFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        bFirstTappedConnectedFeature = true;
                    }
                    else
                    {
                        #region Snapping Error commented
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "LTCable";
                        pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        //_bSnappingErrorCheck = true;                                                                                                                      
                        #endregion
                    }
                }
                #region Code written on 17-May-2010 //To tackle cases where NetJunctions are not found at intersection
                else
                {
                    string sLTCableOID = pCurrentLTCableFeature.OID.ToString();
                    IFeature pLTCable = Setting_Layers_Path.getFeature4GenericCollection(_pFirstPoint, 
                                                                                pLTCableFeatureLayer, sLTCableOID);
                    //IFeature pLTCable = getFeature4GenericCollection(pCurrentLTCableFeature, pLTCableFeatureLayer,
                    //                                                                                    sLTCableOID);
                    if (pLTCable != null)
                    {
                        pRow["SrNo"] = "";
                        pRow["End1OID"] = pLTCable.OID.ToString();
                        pRow["CurrentFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        bFirstTappedConnectedFeature = true;
                    }
                    else
                    {
                        #region Snapping Error commented
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "LTCable";
                        pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        //_bSnappingErrorCheck = true;                                                                                                                      
                        #endregion
                    }
                }
                #endregion
            }

            //Second Point Feature
            pSPFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pLastPoint, 0.014,
                                                                                pServicePointFeatureLayer, "");
            if (pSPFeature != null)
            {
                string strCOID = pSPFeature.get_Value(iCOID).ToString();
                pRow["COID"] = strCOID;
                //pRow["UnitConsumption"] = get_LOAD(strCOID);
                pRow["UnitConsumption"] = "10";
                //bFirstTappedConnectedFeature=true;
                bSecondTappedConnectedFeature = true;
            }
            else
            {
                IFeature pNetJunctFeature = Setting_Layers_Path.GetFirstBufferedFeature(_pLastPoint, 0.014,
                                                                                pNetJunctionFeatureLayer, "");
                if (pNetJunctFeature != null)
                {
                    string sLTCableOID = pCurrentLTCableFeature.OID.ToString();
                    IFeature pLTCable = getFeature4GenericCollection(pNetJunctFeature, pLTCableFeatureLayer,
                                                                                                    sLTCableOID);
                    if (pLTCable != null)
                    {
                        pRow["SrNo"] = "";
                        pRow["End1OID"] = pLTCable.OID.ToString();
                        pRow["CurrentFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        bSecondTappedConnectedFeature = true;
                    }
                    else
                    {
                        #region Snapping Error commented
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "LTCable";
                        pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        //_bSnappingErrorCheck = true;                                                                                                                      
                        #endregion
                    }
                }
                #region To tackle cases where NetJunctions are not found at intersection
                else
                {
                    string sLTCableOID = pCurrentLTCableFeature.OID.ToString();
                    IFeature pLTCable = Setting_Layers_Path.getFeature4GenericCollection(_pLastPoint, 
                                                                                pLTCableFeatureLayer, sLTCableOID);
                    //IFeature pLTCable = Setting_Layers_Path.GetFirstBufferedFeature(_pLastPoint, 0.014,
                                                                                            //pLTCableFeatureLayer, "");
                    //IFeature pLTCable = getFeature4GenericCollection(pCurrentLTCableFeature, pLTCableFeatureLayer,
                    //                                                                                    sLTCableOID);
                    if (pLTCable != null)
                    {
                        pRow["SrNo"] = "";
                        pRow["End1OID"] = pLTCable.OID.ToString();
                        pRow["CurrentFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        bSecondTappedConnectedFeature = true;
                    }
                    else
                    {
                        #region Snapping Error commented
                        iErrorCount++;
                        DataRow pErrorRow = dtErrorResults.NewRow();
                        pErrorRow["SrNo"] = iErrorCount.ToString();
                        pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                        pErrorRow["LayerName"] = "LTCable";
                        pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                        pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                        dtErrorResults.Rows.Add(pErrorRow);
                        //_bSnappingErrorCheck = true;                                                                                                                      
                        #endregion
                    }
                }
                #endregion
            }
            //If the both ends of tapped feature are true
            if (bFirstTappedConnectedFeature && bSecondTappedConnectedFeature)
            {
                bTappedConnectedFeature = true;
            }
            return bTappedConnectedFeature;
        }
        
        #endregion      


        #region For Tackling FUMP pillars
        //Changes in the code made on 6-May-10
        #region Commented GetFUMPDetails (code before 06-May-10)
        //private void GetFUMPDetails(IFeature pBusBarLV, ref DataRow _pRow, int iEnd)
        //{
        //    int iflag_BB = 1;
        //    int iFlag = 1;

        //    #region Checking for FUMP Pillar
        //    GenericCollection pFuse_FeatureGenericCollection = new GenericCollection();
        //    CommonGISFunc.AllSpatialQueryFeatures(pBusBarLV.Shape, pFuseFeatureLayer,
        //                                            esriSpatialRelEnum.esriSpatialRelIntersects, "",
        //                                            ref pFuse_FeatureGenericCollection);

        //    string sFusePosition = "";
        //    foreach (IFeature pFuseFeat in pFuse_FeatureGenericCollection)
        //    {
        //        sFusePosition += pFuseFeat.get_Value(iFusePositionFldInd).ToString();
        //    }

        //    bool sFusepositionflag = true;
        //    if (sFusePosition.Length > 0)
        //        sFusepositionflag = sFusePosition.Contains("0");
        //    if (pFuse_FeatureGenericCollection.Count == 2 && sFusepositionflag == false)
        //    {
        //        #region For FUMP whose 2 fuse are closed
        //        ArrayList arrBB_OIDS = new ArrayList();
        //        arrBB_OIDS.Add(pBusBarLV.OID.ToString());
        //        IFeature pTempBBFeat = null;
        //        foreach (IFeature pFuseFeat in pFuse_FeatureGenericCollection)
        //        {
        //            pTempBBFeat = CommonGISFunc.FirstSpatialQueryFeature(pFuseFeat.Shape, pBusBarFeatureLayer,
        //                                                                      esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                                      "");
        //            if (pTempBBFeat != null)
        //            {
        //                if (arrBB_OIDS.Contains(pTempBBFeat.OID.ToString()))
        //                {
        //                }
        //                else
        //                    arrBB_OIDS.Add(pTempBBFeat.OID.ToString());
        //            }
        //            pTempBBFeat = null;
        //        }
        //        if (arrBB_OIDS.Count == 3)
        //        {

        //            IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBusBarLV.Shape, pPillarFeatureLayer,
        //                                                                             esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                                                          "");
        //            if (!arrUniPillarID.Contains(pPillarFeature.OID.ToString()))
        //            {
        //                arrUniPillarID.Add(pPillarFeature.OID.ToString());
        //                DataRow pRow = dtPillarTable.NewRow();
        //                pRow["PillarOID"] = pPillarFeature.OID.ToString();
        //                pRow["BusBarLV-I"] = arrBB_OIDS[0].ToString();
        //                pRow["BusBarLV-II"] = arrBB_OIDS[1].ToString();
        //                pRow["BusBarLV-III"] = arrBB_OIDS[2].ToString();
        //                dtPillarTable.Rows.Add(pRow);
        //            }
        //        }
        //        #endregion
        //    }
        //    else
        //    {
        //        IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
        //        IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
        //        IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
        //        IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
        //        string sModifiedBBLVOID = "";
        //        #region For _pFirst Point Feature
        //        IFeature pFuseFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)pFirstPoint,
        //                                                                        pFuseFeatureLayer,
        //                                                                        esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                                        "");
        //        if (pFuseFeature != null)
        //        {
        //            iflag_BB = 1;
        //            iFlag = 1;
        //            if (pFuseFeature.get_Value(iFusePositionFldInd).ToString() == "1")
        //            {
        //                // setting BBOID
        //                sModifiedBBLVOID = setBusBarLVOID4rFUMP(pFuseFeature, pBusBarLV, iflag_BB);
        //                _pRow["End" + iEnd.ToString() + "BusBarLV"] = sModifiedBBLVOID;                      
        //            }                   
        //        }               
        //        #endregion

        //        pFuseFeature = null;

        //        #region For _pLast Point Feature
        //        pFuseFeature = CommonGISFunc.FirstSpatialQueryFeature((IGeometry)pLastPoint,
        //                                                                       pFuseFeatureLayer,
        //                                                                       esriSpatialRelEnum.esriSpatialRelIntersects,
        //                                                                       "");
        //        if (pFuseFeature != null)
        //        {
        //           // iflag_BB = 1;
        //            iFlag = 1;
        //            if (pFuseFeature.get_Value(iFusePositionFldInd).ToString() == "1")
        //            {                        
        //                iflag_BB = 0;
        //                sModifiedBBLVOID = setBusBarLVOID4rFUMP(pFuseFeature, pBusBarLV, iflag_BB);
        //                _pRow["End" + iEnd.ToString() + "BusBarLV"] = sModifiedBBLVOID;
        //            }
        //        }           
        //        #endregion              
        //    }
        //    #endregion       
        //}
        #endregion                 

        private string setBusBarLVOID4rFUMP(IFeature pFuseFeature, IFeature pBBFeature, int _iFlag_BB)
        {
            //string strLabelling = "";
            string sreturn = null;
            try
            {
                string sBBOID = pBBFeature.OID.ToString();
                IFeature p2BBLV = getFeature4GenericCollection(pFuseFeature, pBusBarFeatureLayer, sBBOID);

                if (p2BBLV != null)
                {
                    string str2BBLV = p2BBLV.OID.ToString();

                    if (_iFlag_BB == 0)
                        sreturn = str2BBLV + "_" + sBBOID;
                    else
                        sreturn = sBBOID + "_" + str2BBLV;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region

                #endregion
            }
            return sreturn;
        }
        #endregion

        //DataTable Queries
        #region DataTable related Queries       

        private void FillErrorTable(string sErrorType, string sLayerName, string sOID)
        {
            iErrorCount++;
            DataRow pErrorRow = dtErrorResults.NewRow();
            pErrorRow["SrNo"] = iErrorCount.ToString();
            pErrorRow["LayerName"] = sLayerName;
            pErrorRow["ErrorType"] = sErrorType;
            pErrorRow["OID"] = sOID;
            dtErrorResults.Rows.Add(pErrorRow); 
        }

        #endregion

        #region For Tackling LTCables that are tapped at one end but not connected to SP at the other

        private bool GetLTCableWithoutTapping(IPoint _pFirstPoint, IPoint _pLastPoint, IFeature _pFeature, ref DataRow _oConnectedLTCableWithoutTappingRow)
        {
            bool bReturn = false;
            try
            {
                bool bFirstConnectedFeature = getConnectedFeature(_pFirstPoint, _pFeature, ref _oConnectedLTCableWithoutTappingRow, "1");
                bool bSecondConnectedFeature = getConnectedFeature(_pLastPoint, _pFeature, ref _oConnectedLTCableWithoutTappingRow, "2");

                if (bFirstConnectedFeature && bSecondConnectedFeature)
                {
                    bReturn = true;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
            return bReturn;
        }

        private bool getConnectedFeature(IPoint _pPoint, IFeature _pFeature, ref DataRow pRow, string _strEnd)
        {
            bool bGetConnectedFeature = false;
            try
            {
                //If connected Feature is BusDropper                                
                IFeature pBDFeature = Setting_Layers_Path.GetFirstBufferedFeature((IGeometry)_pPoint, 0.00101,
                                                                        pBusBarFeatureLayer, "SUBTYPECD = 4");
                if (pBDFeature != null)
                {
                    #region  Chking & getting Fuse details
                    IFeature pFuse = getIntersectingFeature(pBDFeature, pFuseFeatureLayer, "");
                    if (pFuse != null)
                    {
                        pRow["End" + _strEnd + "OID"] = pFuse.OID.ToString();
                        pRow["End" + _strEnd + "Type"] = pFuse.Class.ObjectClassID.ToString();
                        pRow["End" + _strEnd + "No"] = pFuse.OID.ToString();
                        pRow["End" + _strEnd + "Position"] = pFuse.get_Value(iFusePositionFldInd).ToString();
                        getPillarFeature(pFuse, ref pRow, Convert.ToInt32(_strEnd));
                        pBDFeature = getBusDropperFeature(pFuse, pBDFeature.OID.ToString());
                        #region Chking & getting BD details
                        if (pBDFeature != null)
                        {
                            IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.00101,
                                                                                pBusBarFeatureLayer, "SUBTYPECD = 3");
                            #region Chking & getting BC details
                            if (pBusBarLV != null)
                            {
                                pRow["End" + _strEnd + "BusBarLV"] = pBusBarLV.OID.ToString();
                                IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                                IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                                IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                                IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                                BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, Convert.ToInt32(_strEnd), pBusBarLV);                                
                                bGetConnectedFeature = true;                                
                            }
                            else
                            {
                                #region Snapping Error
                                //iErrorCount++;
                                //DataRow pErrorRow = dtErrorResults.NewRow();
                                //pErrorRow["SrNo"] = iErrorCount.ToString();
                                //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                                //pErrorRow["LayerName"] = "BusDropper";
                                //pErrorRow["ErrorType"] = "No Feature found connected to BusDropper";
                                //pErrorRow["OID"] = pBDFeature.OID.ToString();
                                //dtErrorResults.Rows.Add(pErrorRow);
                                //_bSnappingErrorCheck = true;
                                #endregion
                            }
                            #endregion  Chking & getting BC details
                        }
                        else
                        {
                            #region Snapping Error
                            //iErrorCount++;
                            //DataRow pErrorRow = dtErrorResults.NewRow();
                            //pErrorRow["SrNo"] = iErrorCount.ToString();
                            //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                            //pErrorRow["LayerName"] = "Fuse";
                            //pErrorRow["ErrorType"] = "No Feature found connected to Fuse";
                            //pErrorRow["OID"] = pFuse.OID.ToString();
                            //dtErrorResults.Rows.Add(pErrorRow);
                            //_bSnappingErrorCheck = true;
                            #endregion
                        }
                        #endregion  Chking & getting BD details
                    }
                    else
                    {
                        IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.00101,
                                                                              pBusBarFeatureLayer, "SUBTYPECD = 3");
                        #region Chking & getting BC details
                        if (pBusBarLV != null)
                        {
                            pRow["End" + _strEnd + "OID"] = pBDFeature.OID.ToString();
                            pRow["End" + _strEnd + "Type"] = pBDFeature.Class.ObjectClassID.ToString();
                            pRow["End" + _strEnd + "No"] = pBDFeature.OID.ToString();
                            pRow["End" + _strEnd + "Position"] = "1";
                            getPillarFeature(pBusBarLV, ref pRow, Convert.ToInt32(_strEnd));

                            pRow["End" + _strEnd + "BusBarLV"] = pBusBarLV.OID.ToString();
                            IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                            IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                            IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                            IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);

                            BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, Convert.ToInt32(_strEnd), pBusBarLV);
                            //   pRow["COID"] = "0";
                            //   pRow["UnitConsumption"] = "0";
                            bGetConnectedFeature = true;
                        }
                        else
                        {
                            #region Snapping Error
                            //iErrorCount++;
                            //DataRow pErrorRow = dtErrorResults.NewRow();
                            //pErrorRow["SrNo"] = iErrorCount.ToString();
                            //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                            //pErrorRow["LayerName"] = "LTCable";
                            //pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                            //pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                            //dtErrorResults.Rows.Add(pErrorRow);
                            //_bSnappingErrorCheck = true;
                            #endregion
                        }
                        #endregion  Chking & getting BC details
                    }
                    #endregion  Chking & getting Fuse details
                }
                else
                {
                    #region If connected Feature is LTCable
                    string strOID = _pFeature.OID.ToString();
                    IFeature pLTCableFeature = Setting_Layers_Path.getFeature4GenericCollection(_pPoint, pLTCableFeatureLayer, strOID);
                    if (pLTCableFeature != null)
                    {
                        #region If connected to LTCable
                        pBDFeature = CommonGISFunc.FirstSpatialQueryFeature(pLTCableFeature.Shape,
                                                                       pBusBarFeatureLayer,
                                                                       esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                       "SUBTYPECD = 4");
                        if (pBDFeature != null)
                        {
                            #region  Chking & getting Fuse details
                            IFeature pFuse = getIntersectingFeature(pBDFeature, pFuseFeatureLayer, "");
                            if (pFuse != null)
                            {
                                pRow["End" + _strEnd + "OID"] = pFuse.OID.ToString();
                                pRow["End" + _strEnd + "Type"] = pFuse.Class.ObjectClassID.ToString();
                                pRow["End" + _strEnd + "No"] = pFuse.OID.ToString();
                                pRow["End" + _strEnd + "Position"] = pFuse.get_Value(iFusePositionFldInd).ToString();
                                getPillarFeature(pFuse, ref pRow, Convert.ToInt32(_strEnd));
                                pBDFeature = getBusDropperFeature(pFuse, pBDFeature.OID.ToString());
                                #region Chking & getting BD details
                                if (pBDFeature != null)
                                {
                                    IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.00101,
                                                                                        pBusBarFeatureLayer, "SUBTYPECD = 3");
                                    #region Chking & getting BC details
                                    if (pBusBarLV != null)
                                    {
                                        pRow["End" + _strEnd + "BusBarLV"] = pBusBarLV.OID.ToString();
                                        IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                                        IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                                        IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                                        IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                                        BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, Convert.ToInt32(_strEnd), pBusBarLV);
                                        bGetConnectedFeature = true;
                                    }
                                    else
                                    {
                                        #region Snapping Error
                                        //iErrorCount++;
                                        //DataRow pErrorRow = dtErrorResults.NewRow();
                                        //pErrorRow["SrNo"] = iErrorCount.ToString();
                                        //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                                        //pErrorRow["LayerName"] = "BusDropper";
                                        //pErrorRow["ErrorType"] = "No Feature found connected to BusDropper";
                                        //pErrorRow["OID"] = pBDFeature.OID.ToString();
                                        //dtErrorResults.Rows.Add(pErrorRow);
                                        //_bSnappingErrorCheck = true;
                                        #endregion
                                    }
                                    #endregion  Chking & getting BC details
                                }
                                else
                                {
                                    #region Snapping Error
                                    //iErrorCount++;
                                    //DataRow pErrorRow = dtErrorResults.NewRow();
                                    //pErrorRow["SrNo"] = iErrorCount.ToString();
                                    //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                                    //pErrorRow["LayerName"] = "Fuse";
                                    //pErrorRow["ErrorType"] = "No Feature found connected to Fuse";
                                    //pErrorRow["OID"] = pFuse.OID.ToString();
                                    //dtErrorResults.Rows.Add(pErrorRow);
                                    //_bSnappingErrorCheck = true;
                                    #endregion
                                }
                                #endregion  Chking & getting BD details
                            }
                            else
                            {
                                IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(pBDFeature.Shape, 0.00101,
                                                                                      pBusBarFeatureLayer, "SUBTYPECD = 3");
                                #region Chking & getting BC details
                                if (pBusBarLV != null)
                                {
                                    pRow["End" + _strEnd + "OID"] = pBDFeature.OID.ToString();
                                    pRow["End" + _strEnd + "Type"] = pBDFeature.Class.ObjectClassID.ToString();
                                    pRow["End" + _strEnd + "No"] = pBDFeature.OID.ToString();
                                    pRow["End" + _strEnd + "Position"] = "1";
                                    getPillarFeature(pBusBarLV, ref pRow, Convert.ToInt32(_strEnd));

                                    pRow["End" + _strEnd + "BusBarLV"] = pBusBarLV.OID.ToString();
                                    IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                                    IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                                    IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                                    IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);

                                    BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, Convert.ToInt32(_strEnd), pBusBarLV);
                                    //   pRow["COID"] = "0";
                                    //   pRow["UnitConsumption"] = "0";
                                    bGetConnectedFeature = true;
                                }
                                else
                                {
                                    #region Snapping Error
                                    //iErrorCount++;
                                    //DataRow pErrorRow = dtErrorResults.NewRow();
                                    //pErrorRow["SrNo"] = iErrorCount.ToString();
                                    //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                                    //pErrorRow["LayerName"] = "LTCable";
                                    //pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                                    //pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                                    //dtErrorResults.Rows.Add(pErrorRow);
                                    //_bSnappingErrorCheck = true;
                                    #endregion
                                }
                                #endregion  Chking & getting BC details
                            }
                            #endregion  Chking & getting Fuse details
                        }
                        #endregion
                    }
                    #endregion
                    else
                    {
                        #region Connected to BusBar-LV
                        //IFeature pBusBarLV = getIntersectingFeature(_pPoint, pBusBarFeatureLayer,
                        //"SUBTYPECD = 3");
                        IFeature pBusBarLV = Setting_Layers_Path.GetFirstBufferedFeature(_pPoint, 0.00101,
                                                                               pBusBarFeatureLayer, "SUBTYPECD = 3");
                        if (pBusBarLV != null)
                        {
                            #region Chk for FUMP
                            IFeature pPillarFeature = CommonGISFunc.FirstSpatialQueryFeature(pBusBarLV.Shape,
                                                                               pPillarFeatureLayer,
                                                                               esriSpatialRelEnum.esriSpatialRelIntersects,
                                                                               "");
                            int iFUMPChk = 0;
                            if (pPillarFeature != null)
                            {

                                string strSubType = CommonGISFunc.GetFieldValueFN(pPillarFeature, "SUBTYPECD", true);
                                if (strSubType == "4")
                                {
                                    iFUMPChk = 1;
                                }
                            }
                            if (iFUMPChk == 1)
                            {
                                GetFUMPDetails(pBusBarLV, ref pRow, Convert.ToInt32(_strEnd));
                            }
                            #endregion
                            else
                            {
                                pRow["End" + _strEnd + "OID"] = pBusBarLV.OID.ToString();
                                pRow["End" + _strEnd + "Type"] = pBusBarLV.Class.ObjectClassID.ToString();
                                pRow["End" + _strEnd + "No"] = pBusBarLV.OID.ToString();
                                pRow["End" + _strEnd + "Position"] = "1";
                                pRow["End" + _strEnd + "BusBarLV"] = pBusBarLV.OID.ToString();
                                IPolyline pBBLVPolyLine = (IPolyline)pBusBarLV.Shape;
                                IPointCollection pBBLVPointCollection = (IPointCollection)pBBLVPolyLine;
                                IPoint pFirstPoint = pBBLVPointCollection.get_Point(0);
                                IPoint pLastPoint = pBBLVPointCollection.get_Point(pBBLVPointCollection.PointCount - 1);
                                BusCouplerdetails(pFirstPoint, pLastPoint, ref pRow, Convert.ToInt32(_strEnd), pBusBarLV);
                                getPillarFeature(pBusBarLV, ref pRow, Convert.ToInt32(_strEnd));
                                bGetConnectedFeature = true;
                            }
                        }
                        else
                        {
                            //error table entry for no BusBarLV 
                            #region Snapping Error commented
                            //iErrorCount++;
                            //DataRow pErrorRow = dtErrorResults.NewRow();
                            //pErrorRow["SrNo"] = iErrorCount.ToString();
                            //pErrorRow["LTCableFeatureOID"] = pCurrentLTCableFeature.OID.ToString();
                            //pErrorRow["LayerName"] = "LTCable";
                            //pErrorRow["ErrorType"] = "No Feature found connected to LTCable";
                            //pErrorRow["OID"] = pCurrentLTCableFeature.OID.ToString();
                            //dtErrorResults.Rows.Add(pErrorRow);
                            //_bSnappingErrorCheck = true;                                                                                                                      
                            #endregion
                        }
                        #endregion
                    }
                }
                return bGetConnectedFeature;
            }
            catch (Exception err)
            {
                MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source +
                                                 "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
            }
            finally
            {
                #region
                #endregion
            }
            return bGetConnectedFeature;
        }
        #endregion
    }
}

#region getEnd2Details & get_FirstConnectedFeature commented
//private bool get_FirstConnectedFeature(ESRI.ArcGIS.Geometry.IPoint pFirstPoint, ref DataRow pRow, string Edge, IFeature _LTEdge)
//{
//    bool bReturn = false;

//    try
//    {
//    }
//    catch (Exception err)
//    {
//        MessageBox.Show("Message : " + err.Message + "\nSource: " + err.Source + "\nStack: " + err.StackTrace + "\nTargetSite: " + err.TargetSite);
//    }
//    return bReturn;
//}
//
#endregion      

