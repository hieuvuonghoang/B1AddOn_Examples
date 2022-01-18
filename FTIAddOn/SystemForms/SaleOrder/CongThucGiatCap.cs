using FTIB1Core.Models.SAPObject;
using FTIGlobal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using FTIB1Core.Models;

namespace FTIAddOn.SystemForms.SaleOrder
{
    public class CongThucGiatCap
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        private const string MATRIX_UID = "38";
        private const string UDT1_UID = "UDT1";
        private const string UDT2_UID = "UDT2";
        private const string UDT3_UID = "UDT3";
        private SAPbouiCOM.DataTable uDT1 =>
            oForm.DataSources.DataTables.Item(UDT1_UID);
        private SAPbouiCOM.UserDataSource uDT2 =>
            oForm.DataSources.UserDataSources.Item(UDT2_UID);
        private SAPbouiCOM.UserDataSource uDT3 =>
            oForm.DataSources.UserDataSources.Item(UDT3_UID);

        public CongThucGiatCap(string formUID, SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            this.SBO_Application = SBO_Application;
            this.oCompany = oCompany;
            oForm = this.SBO_Application.Forms.Item(formUID);
            oMatrix = oForm.Items.Item(MATRIX_UID).Specific;
        }
        
        public void MenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            switch (pVal.MenuUID)
            {
                case "1282":
                    AddNewSaleOrder(ref pVal, out bubbleEvent);
                    break;
                case "1293":
                    MatrixDeleteRow(ref pVal, out bubbleEvent);
                    break;
            }
        }

        private void AddNewSaleOrder(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if(!pVal.BeforeAction)
            {
                uDT1.Rows.Clear();
            }
        }

        public void FormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            switch (businessObjectInfo.EventType)
            {
                case BoEventTypes.et_FORM_DATA_LOAD:
                    FormDataLoad(ref businessObjectInfo, out bubbleEvent);
                    break;
            }
        }
        public void ItemEvent(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            switch (pVal.EventType)
            {
                case BoEventTypes.et_FORM_LOAD:
                    FormLoad(ref pVal, out bubbleEvent);
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    LostFocus(ref pVal, out bubbleEvent);
                    break;
            }
        }
        public void RightClick(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            switch (eventInfo.ItemUID)
            {
                case MATRIX_UID:
                    GetRowMatrixRightClick(ref eventInfo, out BubbleEvent);
                    break;
            }
        }

        private void FormLoad(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            AddDataSource(ref pVal, out bubbleEvent);
        }
        private void AddDataSource(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (!pVal.BeforeAction)
            {
                var uDT1 = oForm.DataSources.DataTables.Add(UDT1_UID);
                uDT1.Columns.Add("1", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                uDT1.Columns.Add("110", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.UserDataSources.Add(UDT2_UID, BoDataType.dt_LONG_NUMBER);
                oForm.DataSources.UserDataSources.Add(UDT3_UID, BoDataType.dt_LONG_NUMBER);
            }
        }
        private void GetRowMatrixRightClick(ref ContextMenuInfo eventInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (eventInfo.BeforeAction)
            {
                uDT2.Value = eventInfo.Row.ToString();
            }
        }
        private void FormDataLoad(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            SetDataUDT1(ref businessObjectInfo, out bubbleEvent);
        }
        private void SetDataUDT1(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (!businessObjectInfo.BeforeAction)
            {
                uDT1.Rows.Clear();
                var netMatrix = PublicFunctions.GetNetMatrix(oMatrix);
                var oMatrixColumnsInfors = PublicFunctions.GetColumnsInforMatrixs(netMatrix);
                var indexColumn = 0;
                if (netMatrix.Rows.Length > 0)
                {
                    var itemCode = "";
                    MatrixRow rowMatrix = null;
                    for (var i = 0; i < netMatrix.Rows.Length; i++)
                    {
                        rowMatrix = netMatrix.Rows[i];
                        indexColumn = oMatrixColumnsInfors["1"].IndexCloumn;
                        itemCode = rowMatrix.Columns[indexColumn].Value;
                        if (string.IsNullOrEmpty(itemCode))
                            continue;
                        uDT1.Rows.Add();
                        uDT1.SetValue("1", i, itemCode);
                        indexColumn = oMatrixColumnsInfors["110"].IndexCloumn;
                        uDT1.SetValue("110", i, rowMatrix.Columns[indexColumn].Value);
                    }
                }
            }
        }
        private RDR1 GetRowData(int index)
        {
            var rDR1 = new RDR1();
            var netMatrix = PublicFunctions.GetNetMatrix(oMatrix);
            var oMatrixColumnsInfors = PublicFunctions.GetColumnsInforMatrixs(netMatrix);
            var rowMatrix = netMatrix.Rows[index - 1];
            var indexColumn = 0;
            indexColumn = oMatrixColumnsInfors["1"].IndexCloumn;
            rDR1.ItemCode = rowMatrix.Columns[indexColumn].Value;
            indexColumn = oMatrixColumnsInfors["3"].IndexCloumn;
            rDR1.Dscription = rowMatrix.Columns[indexColumn].Value;
            indexColumn = oMatrixColumnsInfors["15"].IndexCloumn;
            rDR1.DiscPrcnt = string.IsNullOrEmpty(rowMatrix.Columns[indexColumn].Value) ?
                0 : decimal.Parse(rowMatrix.Columns[indexColumn].Value);
            indexColumn = oMatrixColumnsInfors["110"].IndexCloumn;
            rDR1.LineNum = string.IsNullOrEmpty(rowMatrix.Columns[indexColumn].Value) ?
                0 : int.Parse(rowMatrix.Columns[indexColumn].Value);
            return rDR1;
        }
        private Dictionary<int, RDR1> GetDataUDT1()
        {
            var uDT1 = oForm.DataSources.DataTables.Item(UDT1_UID);
            var rDR1s = new List<RDR1>();
            var itemCode = "";
            for (var i = 0; i < uDT1.Rows.Count; i++)
            {
                itemCode = uDT1.GetValue("1", i);
                if (string.IsNullOrEmpty(itemCode))
                    continue;
                var rDR1Tmp = new RDR1()
                {
                    ItemCode = uDT1.GetValue("1", i),
                    LineNum = int.Parse(uDT1.GetValue("110", i)),
                    Index = i,
                };
                rDR1s.Add(rDR1Tmp);
            }
            var dI_RDR1s = rDR1s.ToDictionary(it => it.LineNum, it => it);
            return dI_RDR1s;
        }
        private void LostFocus(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            switch (pVal.ItemUID)
            {
                case MATRIX_UID:
                    MatrixLostFocus(ref pVal, out bubbleEvent);
                    break;
            }
        }
        private void MatrixLostFocus(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            switch (pVal.ColUID)
            {
                case "1"://ItemCode
                case "3"://Dscriptions
                    FillTyLeHoaHong(ref pVal, out bubbleEvent);
                    break;
            }
        }
        private void FillTyLeHoaHong(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (!pVal.BeforeAction)
            {
                var rDR1 = GetRowData(pVal.Row);
                if (string.IsNullOrEmpty(rDR1.ItemCode))
                {
                    SBO_Application.SetStatusBarMessage("ItemCode IsNullOrEmpty", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    return;
                }
                var dI_RDR1s = GetDataUDT1();
                var isAddLine = true;
                if (dI_RDR1s.ContainsKey(rDR1.LineNum))
                {
                    var dI_RDR1 = dI_RDR1s[rDR1.LineNum];
                    if (dI_RDR1.ItemCode == rDR1.ItemCode)
                    {
                        isAddLine = false;
                    }
                    else
                    {
                        uDT1.SetValue("1", dI_RDR1.Index, rDR1.ItemCode);
                    }
                }
                else
                {
                    uDT1.Rows.Add();
                    rDR1.Index = uDT1.Rows.Count - 1;
                    uDT1.SetValue("1", rDR1.Index, rDR1.ItemCode);
                    uDT1.SetValue("110", rDR1.Index, rDR1.LineNum);
                }
                if (isAddLine)
                {
                    SBO_Application.SetStatusBarMessage("Add Line", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    oMatrix.SetCellWithoutValidation(pVal.Row, "U_CK_CD", rDR1.DiscPrcnt.ToString());
                    //oMatrix.SetCellWithoutValidation(pVal.Row, "U_TCK", "100000");
                }
                else
                {
                    SBO_Application.SetStatusBarMessage("Update Line", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
        }
        private void MatrixDeleteRow(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (pVal.BeforeAction)
            {
                var row = int.Parse(uDT2.Value);
                var rDR1 = GetRowData(row);
                uDT3.Value = rDR1.LineNum.ToString();
            }
            else
            {
                var lineNum = int.Parse(uDT3.Value);
                var dI_RDR1s = GetDataUDT1();
                if (dI_RDR1s.ContainsKey(lineNum))
                {
                    var index = dI_RDR1s[lineNum].Index;
                    uDT1.Rows.Remove(index);
                }
            }
        }
    }
}
