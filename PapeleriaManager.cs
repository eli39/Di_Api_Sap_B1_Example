using Reporting.Service.Core.Clientes;
using System;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Reporting.Service.Core.Papeleria
{
    public class PapeleriaManager : DataRepository
    {
        public void CoreInsertOrderStationerySap()
        {
            int lretcode;
            int nResult;
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            oCompany.CompanyDB = "###";
            oCompany.Server = "###";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish;
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sa";
            oCompany.UserName = "###";
            oCompany.Password = "sasa";
            oCompany.LicenseServer = "###";
            oCompany.Disconnect();
            nResult = oCompany.Connect();
            if (nResult == 0)
            {
                List<Papeleria> DetalleHeader = new List<Papeleria>();
                DbCommand cmd = this.Database.GetStoredProcCommand("spOrdersStationeryHeader");
                cmd.CommandTimeout = 0;
                IDataReader dr = this.Database.ExecuteReader(cmd);
                int Sequence;
                while (dr.Read())
                {
                    //CABECERA                        
                    SAPbobsCOM.Documents oInvoiceDoc = null;
                    oInvoiceDoc = (SAPbobsCOM.Documents)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit));                        
                    oInvoiceDoc.DocDate = DateTime.Now;
                    oInvoiceDoc.Reference2 = "PAPELERIA";
                    Sequence = (Int32)dr["Sequence"];
                    //DETALLE
                    List<Papeleria> Detalle = new List<Papeleria>();
                    DbCommand cmd2 = this.Database.GetStoredProcCommand("spOrdersStationery2");
                    this.Database.AddInParameter(cmd2, "@Folio", DbType.Int32, Sequence);

                    cmd.CommandTimeout = 0;
                    IDataReader dr2 = this.Database.ExecuteReader(cmd2);
                    int Row = 0;
                    while (dr2.Read())
                    {
                        int BatchRow = 0;
                        oInvoiceDoc.Comments = (string)dr2["Comentario"];
                        oInvoiceDoc.Lines.ItemCode = (string)dr2["ItemCode"];
                        oInvoiceDoc.Lines.WarehouseCode = "CEDIS";
                        oInvoiceDoc.Lines.Quantity = (int)dr2["CantidadAprobada"];
                        oInvoiceDoc.Lines.SetCurrentLine(Row);
                        oInvoiceDoc.Lines.BaseType = 0;
                        oInvoiceDoc.Lines.UnitPrice = 0;
                        oInvoiceDoc.Lines.AccountCode = (string)dr2["Departamento"];
                        oInvoiceDoc.Lines.BatchNumbers.SetCurrentLine(BatchRow);
                        oInvoiceDoc.Lines.BatchNumbers.BatchNumber = (string)dr2["Distnumber"];
                        oInvoiceDoc.Lines.BatchNumbers.Quantity = (int)dr2["CantidadAprobada"];
                        oInvoiceDoc.Lines.BatchNumbers.Add();
                        oInvoiceDoc.Lines.Add();
                        Row++;
                    }
                    dr2.Close();
                    lretcode = oInvoiceDoc.Add();
                    if (lretcode != 0)
                    {
                        oCompany.GetLastError(out int lErrCode, out string sErrMsg);
                        //string errocode = oCompany.GetLastErrorDescription();
                        // CoreErrorSapOrderStationery(myresult, errocode);
                    }
                    else
                    {                       
                        string FolioSap = oCompany.GetNewObjectKey();
                        Convert.ToInt16(FolioSap);
                        CoreUpdateOrderStationerySap(Sequence, FolioSap);
                    }
                }
                dr.Close();
            }
        }
        public List<Papeleria> CoreSearchStationery(string ItemName)
        {
            List<Papeleria> Detalle = new List<Papeleria>();
            DbCommand cmd = this.Database.GetStoredProcCommand("spSearchStationery");
            this.Database.AddInParameter(cmd, "@ItemName", DbType.String, ItemName);
            cmd.CommandTimeout = 0;
            IDataReader dr = this.Database.ExecuteReader(cmd);
            while (dr.Read())
            {
                Detalle.Add(new Papeleria
                {
                    ItemCode = (string)dr["ItemCode"].ToString(),
                    ItemName = (string)dr["ItemName"].ToString(),
                    Marca = (string)dr["U_Marca"].ToString(),
                    Stock = DBNull.Value.Equals(dr["Stock"]) ? 0 : (decimal)dr["Stock"]
                });
            }
            return Detalle;
        }
        public bool CoreInsertOrderStationery(string ItemCode, string Comentario, int Cantidad)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spInsertOrderStationery");
            this.Database.AddInParameter(cmd, "@ItemCode", DbType.String, ItemCode);
            this.Database.AddInParameter(cmd, "@Comentario", DbType.String, Comentario);
            this.Database.AddInParameter(cmd, "@Cantidad", DbType.Int32, Cantidad);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }
        public bool CoreInsertFolioPedido(string Departamento, string UsuarioFolio)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spInsertFolioPedido");
            this.Database.AddInParameter(cmd, "@Departamento", DbType.String, Departamento);
            this.Database.AddInParameter(cmd, "@UsuarioFolio", DbType.String, UsuarioFolio);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }
        public List<Papeleria> CoreOrdersStationeryByUser(string UsuarioFolio)
        {
            List<Papeleria> Detalle = new List<Papeleria>();
            DbCommand cmd = this.Database.GetStoredProcCommand("spOrdersStationeryByUser");
            this.Database.AddInParameter(cmd, "@UsuarioFolio", DbType.String, UsuarioFolio);
            cmd.CommandTimeout = 0;
            IDataReader dr = this.Database.ExecuteReader(cmd);
            while (dr.Read())
            {
                Detalle.Add(new Papeleria
                {
                    Sequence = (int)dr["Sequence"],
                    Foliox = (int)dr["Folio"],
                    FolioSap = DBNull.Value.Equals(dr["FolioSap"]) ? " " : (string)dr["FolioSap"],
                    ItemCode = (string)dr["ItemCode"],
                    ItemName = (string)dr["ItemName"],
                    Cantidad = (int)dr["Cantidad"],
                    CantidadAprobada = DBNull.Value.Equals(dr["CantidadAprobada"]) ? 0 : (int)dr["CantidadAprobada"],
                    Comentario = (string)dr["Comentario"].ToString(),
                    ComentarioSistema = DBNull.Value.Equals(dr["ComentarioSistema"]) ? " " : (string)dr["ComentarioSistema"],
                    DistNumber = DBNull.Value.Equals(dr["Distnumber"]) ? " " : (string)dr["Distnumber"],
                    EstatusPedido = (int)dr["EstatusPedido"],
                    FechaPedido = (string)dr["FechaPedido"],
                    CentroCosto = (string)dr["Departamento"],
                    UsuarioFolio = (string)dr["UsuarioFolio"],
                    EstatusFolio = (int)dr["EstatusFolio"]

                });
            }
            return Detalle;
        }
        public List<Papeleria> CoreOrdersStationery()
        {
            List<Papeleria> Detalle = new List<Papeleria>();
            DbCommand cmd = this.Database.GetStoredProcCommand("spOrdersStationery");
            cmd.CommandTimeout = 0;
            IDataReader dr = this.Database.ExecuteReader(cmd);
            while (dr.Read())
            {
                Detalle.Add(new Papeleria
                {
                    Sequence = (int)dr["Sequence"],
                    Foliox = (int)dr["Folio"],
                    FolioSap = DBNull.Value.Equals(dr["FolioSap"]) ? " " : (string)dr["FolioSap"],
                    ItemCode = (string)dr["ItemCode"],
                    ItemName = (string)dr["ItemName"],
                    Cantidad = (int)dr["Cantidad"],
                    CantidadAprobada = DBNull.Value.Equals(dr["CantidadAprobada"]) ? 0 : (int)dr["CantidadAprobada"],
                    Comentario = (string)dr["Comentario"].ToString(),
                    ComentarioSistema = DBNull.Value.Equals(dr["ComentarioSistema"]) ? " " : (string)dr["ComentarioSistema"],
                    DistNumber = DBNull.Value.Equals(dr["Distnumber"]) ? " " : (string)dr["Distnumber"],
                    EstatusPedido = (int)dr["EstatusPedido"],
                    FechaPedido = (string)dr["FechaPedido"],
                    CentroCosto = (string)dr["Departamento"],
                    UsuarioFolio = (string)dr["UsuarioFolio"],
                    EstatusFolio = (int)dr["EstatusFolio"]

                });
            }
            return Detalle;
        }
        public bool CoreAprovetOneOrderStationery(string Sequence, string CantidadAprobada, int Foliox)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spAproveDetailOneOrderStationery");
            this.Database.AddInParameter(cmd, "@Sequence", DbType.String, Sequence);
            this.Database.AddInParameter(cmd, "@CantidadAprobada", DbType.String, CantidadAprobada);
            this.Database.AddInParameter(cmd, "@Folio", DbType.Int32, Foliox);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }
        public bool CoreRejectOneOrderStationery(string Sequence, int Foliox, string ComentarioSistema)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spRejectDetailOneOrderStationery");
            this.Database.AddInParameter(cmd, "@Sequence", DbType.String, Sequence);
            this.Database.AddInParameter(cmd, "@Folio", DbType.Int32, Foliox);
            this.Database.AddInParameter(cmd, "@ComentarioSistema", DbType.String, ComentarioSistema);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }
        public bool CoreUpdateOrderStationerySap(int Sequence, string FolioSap)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spUpdateSapOrderStationery");
            this.Database.AddInParameter(cmd, "@Folio", DbType.String, Sequence);
            this.Database.AddInParameter(cmd, "@FolioSap", DbType.String, FolioSap);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }
        public bool CoreErrorSapOrderStationery(string Sequence, string ComentarioFolio)
        {
            DbCommand cmd = this.Database.GetStoredProcCommand("spErrorSapOrderStationery");
            this.Database.AddInParameter(cmd, "@Sequence", DbType.String, Sequence);
            this.Database.AddInParameter(cmd, "@ComentarioFolio", DbType.String, ComentarioFolio);
            IDataReader dr = this.Database.ExecuteReader(cmd);
            if (dr.RecordsAffected > 0)
                return true;
            else
                return false;
        }

        public List<Papeleria> CoreReportePapeleriaUno(string FecIni, string FecFin)
        {
            List<Papeleria> Detalle = new List<Papeleria>();
            DbCommand cmd = this.Database.GetStoredProcCommand("spReportePapeleriaUno");
            this.Database.AddInParameter(cmd, "@FecIni", DbType.String, FecIni);
            this.Database.AddInParameter(cmd, "@FecFin", DbType.String, FecFin);
            cmd.CommandTimeout = 0;
            IDataReader dr = this.Database.ExecuteReader(cmd);
            while (dr.Read())
            {
                Detalle.Add(new Papeleria
                {
                    ItemCode = (string)dr["SKU"],
                    ItemName = DBNull.Value.Equals(dr["Articulo"]) ? " " : (string)dr["Articulo"],
                    CantidadReporte = (decimal)dr["Cantidad"],                    
                    Total = DBNull.Value.Equals(dr["Total"]) ? 0 : (decimal)dr["Total"]
                });
            }
            return Detalle;
        }
        public List<Papeleria> CoreReportePapeleriaDos(string FecIni, string FecFin)
        {
            List<Papeleria> Detalle = new List<Papeleria>();
            DbCommand cmd = this.Database.GetStoredProcCommand("spReporteCentroCosto");
            this.Database.AddInParameter(cmd, "@FecIni", DbType.String, FecIni);
            this.Database.AddInParameter(cmd, "@FecFin", DbType.String, FecFin);
            cmd.CommandTimeout = 0;
            IDataReader dr = this.Database.ExecuteReader(cmd);
            while (dr.Read())
            {
                Detalle.Add(new Papeleria
                {
                    CentroCosto = DBNull.Value.Equals(dr["CentroCostos"]) ? " " : (string)dr["CentroCostos"],
                    Area = DBNull.Value.Equals(dr["Area"]) ? " " : (string)dr["Area"],
                    Total = DBNull.Value.Equals(dr["Total"]) ? 0 : (decimal)dr["Total"]
                });
            }
            return Detalle;
        }
    }
}