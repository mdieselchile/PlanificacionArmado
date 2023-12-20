using Maestros;
using Maestros.Estados;
using Maestros.OportunidadMejora;
using Maestros.Tipos;
using OrdenTrabajo;
using Planificacion;
using Recursos;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Utilidades;
using static PlanificacionArmado.PlanArmado;
using static System.Net.WebRequestMethods;

namespace PlanificacionArmado
{
    public partial class PlanArmado : Form
    {

        #region Declaraciones
        //Controla que no se ejecuten rutinas causadas por eventos de controles
        private bool EventoCodigo = false;
        private bool RefrescoGrilla = true;
        //Codigo de linea de productos en nivel3, a partir de codigo en gen_nivel3
        protected int CodigoNivel3Homologado = 0;        
        protected string UsuarioConectado = string.Empty;
        protected int RolUsuarioConectado = 0;

        private readonly Color ColorLetra = Color.White;
        private readonly Color ColorAtraso = Color.Orange;
        private readonly Color ColorRetrasado = Color.Yellow;
        private readonly Color ColorSuspendida = Color.CornflowerBlue;
        private readonly Color ColorFueraPlazo = Color.Red;
        private readonly Color ColorAtiempo = Color.Black;
        private readonly Color ColorCargaHHMayorEstima = Color.PaleVioletRed;
        private readonly Color ColorAporteClte = Color.Magenta;
        private const int UnDiaMenosUnMinuto = 1439;
        ///////Constantes
        //5 segundos para mostrar mensajes en barra estado
        private const int INTERVALO_ACTUALIZACION = 5000;
        private const double IMPUESTO = 1.19;        
        private const int ValorMaximoPrioridad = 100;                              
        public const int TeclaBackSpace = 8;

        private const string MsgAviso = "Sólo se pueden actualizar HH estimadas de procesos 'En Ejecución'";
        private const string MsgErrorFechaProd = "Fecha entrega planificada no puede ser menor a la actual";
        private const string MsgErrorFechaRecup = "Fecha entrega recuperación no puede ser menor a la actual";
        private const string MsgOK = "Fecha entrega planificada actualizada";
        private const string MsgOKRecup = "Fecha entrega recuperación actualizada";
        private const string MsgFechaEvalIni = "Fecha inicio evaluación actualizada";
        private const string MsgFechaEvalFin = "Fecha término evaluación actualizada";        
                
        public static MessageBoxIcon MsgIconError = MessageBoxIcon.Error;
        public static MessageBoxIcon MsgIconAviso = MessageBoxIcon.Asterisk;
        public static MessageBoxIcon MsgIconInfo = MessageBoxIcon.Information;
        
        private SortableBindingList<OrdenTrabajo.GestionOrdenTrabajo.AuxGestion> TablaOTs;        
        private List<EstadOrdenTrabajo> ListaEstadosOTTaller;
        
        //Capacidades
        private Capacidades.TipoSeccion TipoSec;
        private decimal MaximoHorasSecciones;
        private const string MsgCarga = "Espere mientras se cargan datos maestros...";
        private string TituloEjeX = "Secciones desarme y armado";
        private string TituloGrafico;
        private bool MostrarOcultarSeguimientos = true;
        #endregion        
        public PlanArmado()
        {
            InitializeComponent();
            Inicializar();            
        }

        public PlanArmado(Vista vista)
        {
            InitializeComponent();
            GetVista = vista;
            Inicializar();
           
        }

        private void Inicializar()
        {
            BotonEjecutar.Click += BotonEjecutar_Click;
            Sucursal.SelectionChangeCommitted += Sucursal_SelectionChangeCommitted;
            Negocio.SelectionChangeCommitted += Negocio_SelectionChangeCommitted;
            Producto.SelectionChangeCommitted += Producto_SelectionChangeCommitted;
            Seccion.SelectionChangeCommitted += Seccion_SelectionChangeCommitted;
            Cliente.SelectionChangeCommitted += Cliente_SelectionChangeCommitted;
            SucursalCliente.SelectionChangeCommitted += SucursalCliente_SelectionChangeCommitted;

            MenuimprimirGraf.Click += MenuimprimirGraf_Click;
            MenuMostrarHH.Click += MenuMostrarHH_Click;
            MenuFecEntrEval.Click += MenuFecEntrEval_Click;
            MenuVentas.Click += MenuVentas_Click;
            MenuAbrirOT.Click += MenuAbrirOT_Click;
            MenuAbrirHojaRuta.Click += MenuAbrirHojaRuta_Click;
            MenuAbrirRutaHoja.Click += MenuAbrirRutaHoja_Click;
            
            MenuActualizarProcParal.Click += MenuActualizarProcParal_Click;

            MenuDetenerTrabajo.Click += MenuDetenerTrabajo_Click;
            
            MenuActualizarFechaInicioCorridos.Click += MenuActualizarFechaInicioCorridos_Click;
            MenuActualizarFechaInicioHabiles.Click += MenuActualizarFechaInicioHabiles_Click;
            MenuActualizarFechaInicioTodosHabiles.Click += MenuActualizarFechaInicioTodosHabiles_Click;
            MenuActualizarFechaInicioTodosCorridos.Click += MenuActualizarFechaInicioTodosCorridos_Click;            
            MenuReanudarTrabajo.Click += MenuReanudarTrabajo_Click;
            MenuActualFechasProcesoHojasCorr.Click += MenuActualFechasProcesoHojasCorr_Click;
            MenuActualFechasProcesoHojas.Click += MenuActualFechasProcesoHojas_Click;            
            MenuActualizarSegProc.Click += MenuActualizarSegProc_Click;
                    
            PlanillaCapacidad.RowEnter += PlanillaCapacidad_RowEnter;                
            PlanillaCapacidadPeriodo.Scroll += PlanillaCapacidadSemanal_Scroll;
            PlanillaCapacidadPeriodo.ColumnWidthChanged += PlanillaCapacidadSemanal_ColumnWidthChanged;
            FechaEntregaProduccionIni.ValueChanged += FechaCapIni_ValueChanged;
            FechaEntregaProduccionFin.ValueChanged += FechaCapFin_ValueChanged;

            PlanillaOT.CellPainting += PlanillaOT_CellPainting;
            PlanillaOT.Paint += PlanillaOT_Paint;
            PlanillaOT.Scroll += PlanillaOT_Scroll;
            PlanillaOT.ColumnWidthChanged += PlanillaOT_ColumnWidthChanged;            
            PlanillaOT.CellClick += PlanillaOT_CellClick;
            PlanillaOT.CellDoubleClick += PlanillaOT_CellDoubleClick;
            PlanillaOT.CellEndEdit += PlanillaOT_CellEndEdit;
            PlanillaOT.CellLeave += PlanillaOT_CellLeave;
            PlanillaOT.CellMouseDown += PlanillaOT_CellMouseDown;
            PlanillaOT.CellValidating += PlanillaOT_CellValidating;
            PlanillaOT.ColumnHeaderMouseClick += PlanillaOT_ColumnHeaderMouseClick;
            PlanillaOT.DataError += PlanillaOT_DataError;
            PlanillaOT.RowEnter += PlanillaOT_RowEnter;
            PlanillaOT.CellContentClick += PlanillaOT_CellContentClick;            

            //PlanillaOT.RowPostPaint += PlanillaOT_RowPostPaint;
            //PlanillaOT.RowPrePaint += PlanillaOT_RowPrePaint;
            //PlanillaOT.CurrentCellChanged += PlanillaOT_CurrentCellChanged;
            //PlanillaOT.RowHeightChanged += PlanillaOT_RowHeightChanged;            
            chkClte.CheckedChanged += ChkClte_CheckedChanged;
            chkSucClte.CheckedChanged += ChkSucClte_CheckedChanged;

            GrillaSegAporte.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegProc.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegMat.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegRepInter.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegRepNac.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegSub.DoubleClick += GrillasSeguimientos_DoubleClick;
            GrillaSegProc.CellFormatting += GrillasSeguimientos_CellFormatting;


            LblSuspend.CheckedChanged += LblSuspend_CheckedChanged;
            LblAtraso.CheckedChanged += LblAtraso_CheckedChanged;
            LblFuera.CheckedChanged += LblFuera_CheckedChanged;
            LblRetraso.CheckedChanged += LblRetraso_CheckedChanged;
            LblHH.CheckedChanged += LblHH_CheckedChanged;
            LblATiempo.CheckedChanged += LblATiempo_CheckedChanged;

            BotonTrabajos.Click += BotonTrabajos_Click;
        }

       
        private void BotonTrabajos_Click(object sender, EventArgs e)
        {
            TrabajosEnCurso();
        }

        #region Enumeraciones

        public enum Vista
        {
            Planificacion, Taller, Comercial,Armado,Recuperacion,Presupuesto,ControlCalidad
        }        
        public Vista GetVista { get; set; }
        private enum CF
        {
            NumeroOT = 0,            
            Sucursal = 1,
            Componente = 2,
            Marca = 3,
            Modelo = 4,
            Cliente = 5,
            Tipo = 6,
            Estado = 7,
            Prioridad = 8,
            Avance = 9,
            HrsEstimadas = 10,
            HrsCargadas = 11,
            HrsPendientes = 12,
           
            FechaRecepcion = 13,
            DiasEnTaller = 14,
            DiasEnProceso=15,
            FechaInicioEval = 16,
            FechaFinEval = 17,
            FechaFinEvalReal = 18,
            FechaReprogEval = 19,
            FechaInicioPlanif = 20,
            FechaFinPlanif = 21,
            FechaFinPlanifReal = 22,
            FechaReprogPlanif = 23,
            FechaInicioPpto = 24,
            FechaFinPpto = 25,
            FechaFinPptoReal = 26,
            FechaReprogPpto = 27,
            FechaInicioRep = 28,
            FechaEntregaRecup = 29,
            FechaInicioArmado = 30,
            FechaFinRep = 31,
            FechaFinRepReal = 32,
            FechaArribo = 33,
            FechaEntregaProd = 34,
            FechaLiberacion = 35,
            FechaDespacho = 36,
            FechaEntregaClte = 37,
            Suspendida = 38,
            CodigoEstado = 39,
            DetalleOT = 40,
            Responsable = 41,
            Asistente = 42,
            Dossier = 43,
            MotivoIncum = 44,
            InfRC = 45,
            InfEval = 46,
            InfFinal = 47,
            Vendedor = 48,
            OcCliente = 49,
            ValorTrabajo = 50,
            OfertaVenta = 51,
            EstadoEntrega = 52,
            CodigoFab = 53,
            AporteClte = 54
        }

        public enum CD
        {
            Imagen,HojaNum,Item,Elemento,Seccion,Proceso, EstadoProc,HorasEstim,
            Dias, HorasCargadas,DiasRestantes,FechaIni,FechaFin,Predecesor,
            CodigoEstadoProc
        }
       
        public enum Area
        {
            Planificacion, RecursosHumanos
        }
        private enum ColCap
        {
            Seccion, HCapacidad, HEstim, HCarga, HDisp, HPorCargar            
        }

        private enum ColTrab
        {
            Trabajador, NumOT, NumHoja, Proceso, FechaIni, FechaFin, HoraIni, HoraFin, TiempoTrans, TiempoEstim
        }
        #endregion

        #region CargaListas
        protected virtual void CargaNivel2()
        {
            try
            {
                var pl = new Planificacion.Planificacion();                
                int codNivel1 = Convert.ToInt16(Sucursal.SelectedValue);                
                string msg = string.Empty;                

                pl.Nivel1 = codNivel1;
                var dt = pl.ListaNivel2(ref msg);
                if (dt != null)
                {                    
                    Negocio.DataSource = dt;
                    Negocio.ValueMember = "Codigo";
                    Negocio.DisplayMember = "Nombre";
                    Negocio.SelectedIndex = -1;
                }
                else
                {                    
                    SetErrorMsg(msg);
                }
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        protected void CargaNivel3()
        {
            try
            {
                var pl = new Planificacion.Planificacion();               
                int codNivel1 = Convert.ToInt16(Sucursal.SelectedValue);                
                int codNivel2 = Convert.ToInt16(Negocio.SelectedValue);
                string msg = string.Empty;               

                pl.Nivel1 = codNivel1;
                pl.Nivel2 = codNivel2;
                var dt = pl.ListaNivel3(ref msg);
                if (dt != null)
                {                    
                    Producto.DataSource = dt;
                    Producto.ValueMember = "Codigo";
                    Producto.DisplayMember = "Nombre";
                    Producto.SelectedIndex = -1;
                }
                else
                {                    
                    SetErrorMsg(msg);
                }
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }            
        }
        protected void CargaNivel5()
        {
            try
            {
                var pl = new Planificacion.Planificacion();                
                int codNivel1 = Convert.ToInt16(Sucursal.SelectedValue);                
                int codNivel3 = Convert.ToInt16(Producto.SelectedValue);
                string msg = string.Empty;
                
                
                pl.Nivel1 = codNivel1;
                pl.Nivel3 = codNivel3;
                CodigoNivel3Homologado = pl.CodigoNivel3DesdeGenNivel3(ref msg);
                pl.Nivel3 = CodigoNivel3Homologado;
                var dt = pl.ListaSecciones(ref msg);
                if (dt != null)
                {                    
                    Seccion.DataSource = dt;
                    Seccion.ValueMember = "Codigo";
                    Seccion.DisplayMember = "Nombre";
                    if (Seccion.Items.Count == 1)
                    {
                        chkSeccion.Checked = true;
                        Seccion.SelectedIndex = 0;
                    }
                    else
                        Seccion.SelectedIndex = -1;
                }
                else
                {                    
                    SetErrorMsg(msg);
                }
               
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }           
        }
        private void CargaNegocios()
        {
            int codNivel1 = Convert.ToInt16(Sucursal.SelectedValue);
            CodigoNivel3Homologado = 0;
            if (Sucursal.SelectedIndex >= 0)
            {                                                    
                var plan = new Planificacion.Planificacion();
                plan.CargaResponsablesOT(ResponsableOT, codNivel1);
                CargaNivel2();
            }
             
            
        }      
        private async Task InicializarObjetos()
        {
            var tarea = new Task(() =>
            {
                CheckForIllegalCrossThreadCalls = false;
                var suc = new Maestros.Organizacion.Nivel1(true);
                suc.Listar(Sucursal);

                var neg = new Maestros.Organizacion.GenNivel2();
                var listaNeg = neg.Listar();
                Negocio.DataSource = listaNeg;
                Negocio.DisplayMember = "Nombre";
                Negocio.ValueMember = "Codigo";
                Negocio.SelectedIndex = -1;

                var prod = new Maestros.Organizacion.GenNivel3();
                var listaProd = prod.Listar();
                Producto.DataSource = listaProd;
                Producto.DisplayMember = "Nombre";
                Producto.ValueMember = "Codigo";
                Producto.SelectedIndex = -1;

                var sec = new Maestros.Organizacion.GenNivel5();
                var lst = sec.ListarProduccion();

                Seccion.DataSource = lst;
                Seccion.DisplayMember = "Nombre";
                Seccion.ValueMember = "Codigo";
                Seccion.SelectedIndex = -1;

                var est = new EstadOrdenTrabajo(1);
                est.Listar(ListaEstados);
                var tipo = new TipoOrdenTrabajo();
                tipo.Listar(ListaTipos);

                var epp = new EstadoProcesoProduccion(1);
                epp.Listar(EstadoProceso);
                var eet = new EstadoEntregaTrabajo(1);
                eet.Listar(EstadoEntrega);

                var cliente = new ClienteProveedor();
                var lista = cliente.Listar("C");
                Cliente.DataSource = lista;
                Cliente.DisplayMember = "RazonSocial";
                Cliente.ValueMember = "RUT";
                Cliente.SelectedIndex = -1;

                var emp = new Empleado(1);
                
                var listaEmp = emp.ListarCarganHora();
                EmpleadoTarea.DataSource = listaEmp;
                EmpleadoTarea.DisplayMember = "Nombre";
                EmpleadoTarea.ValueMember = "Codigo";
                EmpleadoTarea.SelectedIndex = -1;
            });
            tarea.Start();
            await tarea;
        }

        #endregion

        private void OpcionesPorVista()
        {
            switch (GetVista)
            {
                case Vista.Planificacion:
                    Text = "Planificación producción - Vista General";
                    TipoSec = Capacidades.TipoSeccion.Todas;
                    break;
                case Vista.Taller:
                    Text = "Planificación producción - Vista Taller";
                    TipoSec = Capacidades.TipoSeccion.Todas;
                    break;
                case Vista.Comercial:
                    Text = "Planificación producción - Vista Comercial";
                    TipoSec = Capacidades.TipoSeccion.Todas;
                    break;
                case Vista.Armado:
                    Text = "Planificación producción - Vista Desarme y Armado";
                    TipoSec = Capacidades.TipoSeccion.Armado;
                    break;
                case Vista.Recuperacion:
                    Text = "Planificación producción - Vista Fabricación y Recuperación";
                    TipoSec = Capacidades.TipoSeccion.Recuperacion;
                    break;
                case Vista.Presupuesto:
                    Text = "Planificación producción - Vista Presupuesto";
                    TipoSec = Capacidades.TipoSeccion.Todas;
                    break;
                case Vista.ControlCalidad:
                    Text = "Planificación producción - Vista Control de Calidad";
                    TipoSec = Capacidades.TipoSeccion.Todas;
                    break;
                default:
                    break;
            }
        }
        private void EstablecerMesActual()
        {
            FechaEntregaProduccionIni.Value = DateTime.Now;            
            int mesFin = FechaEntregaProduccionFin.Value.Month;
            string fecha2;
            string año = FechaEntregaProduccionIni.Value.Year.ToString();

            int dia = Fechas.DiasDelMes(mesFin);
            if (mesFin.ToString().Length == 1)
                fecha2 = dia.ToString() + "/" + "0" + mesFin.ToString() + "/" + año;
            else
                fecha2 = dia.ToString() + "/" + mesFin.ToString() + "/" + año;

            FechaEntregaProduccionFin.Value = DateTime.Parse(fecha2);

        }
        private void ConfigurarGrillas()
        {
            Grillas.Grilla.FormatoPlanillas(PlanillaOT);
            Grillas.Grilla.EvitaParpadeoGrilla(PlanillaOT);
            Grillas.Grilla.FormatoPlanillas(PlanillaHojas);
            Grillas.Grilla.EvitaParpadeoGrilla(PlanillaHojas);

            Grillas.Grilla.FormatoPlanillas(PlanillaCapacidad);
            Grillas.Grilla.EvitaParpadeoGrilla(PlanillaCapacidad);
            Grillas.Grilla.FormatoPlanillas(PlanillaCapCentros);
            Grillas.Grilla.EvitaParpadeoGrilla(PlanillaCapCentros);
            Grillas.Grilla.FormatoPlanillas(PlanillaCapacidadPeriodo);
            Grillas.Grilla.EvitaParpadeoGrilla(PlanillaCapacidadPeriodo);

            Grillas.Grilla.FormatoPlanillas(GrillaTrabajos);
            Grillas.Grilla.EvitaParpadeoGrilla(GrillaTrabajos);
        }
       
        #region FormatoPlanillaOT
        private void ConfigurarGrillaPlanificacion()
        {
            try
            {
                PlanillaOT.ReadOnly = false;

                Font tipoLetra = new("Arial", 8);
                if (GetVista == Vista.Comercial || GetVista == Vista.Taller || GetVista == Vista.Presupuesto || GetVista == Vista.ControlCalidad)
                {
                    tipoLetra = new Font("Calibri", 10);                    
                    PlanillaOT.ColumnHeadersHeight = 60;
                    PlanillaOT.ColumnHeadersDefaultCellStyle.Font = new Font("Calibri", 10);
                }
                else
                {
                    PlanillaOT.ColumnHeadersHeight = 80;
                }
                PlanillaOT.AllowUserToResizeColumns = true;
                PlanillaOT.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                PlanillaOT.EditMode = DataGridViewEditMode.EditOnKeystroke;
                PlanillaOT.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                PlanillaOT.SelectionMode = DataGridViewSelectionMode.CellSelect;
                
                PlanillaOT.ScrollBars = ScrollBars.Both;                
                PlanillaOT.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                PlanillaOT.Font = tipoLetra;

                var celdaGrid = new DataGridViewCellStyle
                {
                    BackColor = Color.White,                    
                };

                PlanillaOT.DefaultCellStyle = celdaGrid;

                var colSusp = new DataGridViewCheckBoxColumn()
                {
                    Name = "Suspendida_35",
                    HeaderText = "Suspendida"
                };
                var colInfRC = new DataGridViewCheckBoxColumn()
                {
                    Name = "InfRC_43",
                    HeaderText = "IRC",
                    Width = 25
                };
                var colInfEval = new DataGridViewCheckBoxColumn()
                {
                    Name = "InfEval_44",
                    HeaderText = "IE",
                    Width = 25
                };
                var colInfFin = new DataGridViewCheckBoxColumn()
                {
                    Name = "InfFin_45",
                    HeaderText = "IF ",
                    Width = 25
                };
                var colAporte = new DataGridViewCheckBoxColumn()
                {
                    Name = "aporteclte",
                    HeaderText = "Aporte Clte ",
                    Width = 25
                };

                PlanillaOT.Columns.Add("NumeroOT_0", "NumeroOT");
                PlanillaOT.Columns.Add("Sucursal_1", "Sucursal");
                PlanillaOT.Columns.Add("Componente_2", "Componente");
                PlanillaOT.Columns.Add("Marca_3", "Marca");
                PlanillaOT.Columns.Add("Modelo_4", "Modelo");
                PlanillaOT.Columns.Add("Cliente_5", "Cliente");
                PlanillaOT.Columns.Add("Tipo_6", "Tipo");
                PlanillaOT.Columns.Add("Estado_7", "Estado");
                PlanillaOT.Columns.Add("Prioridad_8", "Prioridad");
                PlanillaOT.Columns.Add("Avance_12", "Avance");
                PlanillaOT.Columns.Add("HrsEstimadas_9", "HrsEstimadas");
                PlanillaOT.Columns.Add("HrsCargadas_10", "HrsCargadas");
                PlanillaOT.Columns.Add("HrsPendientes_11", "HrsPendientes");               
                PlanillaOT.Columns.Add("FechaRecepcion_13", "FechaRecepcion");
                PlanillaOT.Columns.Add("DiasEnTaller_14", "DiasEnTaller");
                PlanillaOT.Columns.Add("DiasEnProceso", "DiasEnProceso");
                PlanillaOT.Columns.Add("FechaInicioEval_15", "FechaInicioEval");
                PlanillaOT.Columns.Add("FechaFinEval_16", "FechaFinEval");
                PlanillaOT.Columns.Add("FechaFinEvalReal_17", "FechaFinEvalReal");
                PlanillaOT.Columns.Add("FechaReprogEval", "Fecha Reprog Eval");
                PlanillaOT.Columns.Add("FechaInicioPlanif_18", "FechaInicioPlanif");
                PlanillaOT.Columns.Add("FechaFinPlanif_19", "FechaFinPlanif");
                PlanillaOT.Columns.Add("FechaFinPlanifReal_20", "FechaFinPlanifReal");
                PlanillaOT.Columns.Add("FechaReprogPlanif", "Fecha Reprog Planif");
                PlanillaOT.Columns.Add("FechaInicioPpto_21", "FechaInicioPpto");
                PlanillaOT.Columns.Add("FechaFinPpto_22", "FechaFinPpto");
                PlanillaOT.Columns.Add("FechaFinPptoReal_23", "FechaFinPptoReal");
                PlanillaOT.Columns.Add("FechaReprogPpto", "Fecha Reprog Ppto");
                PlanillaOT.Columns.Add("FechaInicioRep_24", "FechaInicioRep");
                PlanillaOT.Columns.Add("FechaEntregaRecup_25", "FechaEntregaRecup");
                PlanillaOT.Columns.Add("FechaInicioArmado_26", "FechaInicioArmado");
                PlanillaOT.Columns.Add("FechaFinRep_27", "FechaFinRep");
                PlanillaOT.Columns.Add("FechaFinRepReal_28", "FechaFinRepReal");
                PlanillaOT.Columns.Add("FechaArribo_29", "FechaArribo");
                PlanillaOT.Columns.Add("FechaEntregaProd_30", "FechaEntregaProd");
                PlanillaOT.Columns.Add("FechaLiberacion_31", "FechaLiberacion");
                PlanillaOT.Columns.Add("FechaDespacho_32", "FechaDespacho");
                PlanillaOT.Columns.Add("FechaEntregaClte_33", "FechaEntregaClte");                
                PlanillaOT.Columns.Add(colSusp);
                PlanillaOT.Columns.Add("CodigoEstado_36", "CodigoEstado");
                PlanillaOT.Columns.Add("DetalleOT_37", "DetalleOT");
                PlanillaOT.Columns.Add("Responsable_38", "Responsable");
                PlanillaOT.Columns.Add("Asistente_39", "Asistente");
                PlanillaOT.Columns.Add("Dossier_40", "Dossier");
                PlanillaOT.Columns.Add("MotivoIncum_41", "MotivoIncum");

                PlanillaOT.Columns.Add(colInfRC);
                PlanillaOT.Columns.Add(colInfEval);
                PlanillaOT.Columns.Add(colInfFin);

                PlanillaOT.Columns.Add("Vendedor_42", "Vendedor");
                PlanillaOT.Columns.Add("OcCliente_43", "OcCliente");
                PlanillaOT.Columns.Add("ValorTrabajo_44", "ValorTrabajo");
                PlanillaOT.Columns.Add("OfertaVenta_45", "OfertaVenta");
                PlanillaOT.Columns.Add("estadoentrega", "EstadoEntrega");
                PlanillaOT.Columns.Add("codfab", "Codigo fabricado");
                PlanillaOT.Columns.Add(colAporte);

                PlanillaOT.Columns[(int)CF.Cliente].Frozen = true;
                PlanillaOT.Columns[(int)CF.CodigoEstado].Visible = false;
                PlanillaOT.Columns[(int)CF.DetalleOT].Visible = false;
                PlanillaOT.Columns[(int)CF.EstadoEntrega].Visible = false;
                PlanillaOT.Columns[(int)CF.Prioridad].Visible = false;                

                PlanillaOT.Columns[(int)CF.HrsCargadas].DefaultCellStyle.Format = "#,##0.00";
                PlanillaOT.Columns[(int)CF.HrsPendientes].DefaultCellStyle.Format = "#,##0.00";                

                //Alineación derecha columnas númericas
                for (int i = (int)CF.Prioridad; i <= (int)CF.HrsPendientes; i++)
                {
                    PlanillaOT.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                
                //No permitir editar
                for (int i = 0; i < PlanillaOT.ColumnCount; i++)
                {
                    PlanillaOT.Columns[i].ReadOnly = true;
                }
                PlanillaOT.Columns[(int)CF.AporteClte].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.ValorTrabajo].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaOT.Columns[(int)CF.DiasEnProceso].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                PlanillaOT.Columns[(int)CF.DiasEnTaller].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void OcultarColumnasSegunVista()
        {
            //Columnas no visibles si vista dif planificacion
            if (GetVista != Vista.Planificacion)
            {
                TotalVentaPeriodo.Visible = false;
                PlanillaOT.Columns[(int)CF.Tipo].Visible = false;
                //PlanillaOT.Columns[(int)CF.Prioridad].Visible = false;

                PlanillaOT.Columns[(int)CF.HrsEstimadas].Visible = false;
                PlanillaOT.Columns[(int)CF.HrsPendientes].Visible = false;
                PlanillaOT.Columns[(int)CF.Avance].Visible = false;
                PlanillaOT.Columns[(int)CF.HrsCargadas].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaRecepcion].Visible = false;
                PlanillaOT.Columns[(int)CF.DiasEnTaller].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaInicioEval].Visible = false;

                PlanillaOT.Columns[(int)CF.FechaFinEvalReal].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaInicioPlanif].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinPlanif].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinPlanifReal].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaInicioPpto].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinPptoReal].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaInicioRep].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaEntregaRecup].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaInicioArmado].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinRep].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaFinRepReal].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaArribo].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaLiberacion].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaDespacho].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].Visible = false;
                PlanillaOT.Columns[(int)CF.Suspendida].Visible = false;
                PlanillaOT.Columns[(int)CF.Responsable].Visible = false;
                PlanillaOT.Columns[(int)CF.Asistente].Visible = false;
                PlanillaOT.Columns[(int)CF.Dossier].Visible = false;
                PlanillaOT.Columns[(int)CF.MotivoIncum].Visible = false;
                PlanillaOT.Columns[(int)CF.Vendedor].Visible = false;
                PlanillaOT.Columns[(int)CF.OcCliente].Visible = false;
                PlanillaOT.Columns[(int)CF.ValorTrabajo].Visible = false;
                PlanillaOT.Columns[(int)CF.OfertaVenta].Visible = false;
                PlanillaOT.Columns[(int)CF.DiasEnProceso].Visible = false;
                PlanillaOT.Columns[(int)CF.DiasEnTaller].Visible = false;

                PlanillaOT.Columns[(int)CF.InfEval].Visible = false;
                PlanillaOT.Columns[(int)CF.InfRC].Visible = false;
                PlanillaOT.Columns[(int)CF.InfFinal].Visible = false;
            }
            if (GetVista == Vista.Comercial)
            {
                PlanillaOT.Columns[(int)CF.Sucursal].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].Visible = true;
                PlanillaOT.Columns[(int)CF.OcCliente].Visible = true;
                PlanillaOT.Columns[(int)CF.Vendedor].Visible = true;
                PlanillaOT.Columns[(int)CF.DiasEnProceso].Visible = true;
            }
            if (GetVista == Vista.Taller)
            {
                PlanillaOT.Columns[(int)CF.FechaFinRep].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaInicioEval].Visible = true;
                PlanillaOT.Columns[(int)CF.Tipo].Visible = true;
                //PlanillaOT.Columns[(int)CF.Prioridad].Visible = true;
                PlanillaOT.Columns[(int)CF.Asistente].Visible = true;
                PlanillaOT.Columns[(int)CF.Dossier].Visible = true;
                PlanillaOT.Columns[(int)CF.DiasEnTaller].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].Visible = true;
            }
            if (GetVista == Vista.Presupuesto)
            {
                PlanillaOT.Columns[(int)CF.DiasEnTaller].Visible = true;
                PlanillaOT.Columns[(int)CF.Tipo].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaInicioPpto].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaFinPptoReal].Visible = true;

                PlanillaOT.Columns[(int)CF.FechaFinEval].Visible = false;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].Visible = true;
                PlanillaOT.Columns[(int)CF.Responsable].Visible = true;
                PlanillaOT.Columns[(int)CF.Asistente].Visible = true;                
            }
            if (GetVista == Vista.ControlCalidad)
            {                
                PlanillaOT.Columns[(int)CF.Tipo].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaInicioEval].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaFinEval].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaFinEvalReal].Visible = true;                
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].Visible = true;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].Visible = true;
                PlanillaOT.Columns[(int)CF.Responsable].Visible = true;
                PlanillaOT.Columns[(int)CF.Asistente].Visible = true;
                PlanillaOT.Columns[(int)CF.Dossier].Visible = true;
                PlanillaOT.Columns[(int)CF.Vendedor].Visible = true;
            }
        }
        private void ConfigurarGrillaSeguimientos(DataGridView grilla)
        {
            var fuente = new Font(grilla.Font.Name, grilla.Font.Size, FontStyle.Regular);
            grilla.Font = fuente;
            grilla.ScrollBars = ScrollBars.Vertical;
            grilla.AllowUserToResizeColumns = true;
            grilla.RowHeadersVisible = false;
            grilla.AllowUserToAddRows = false;
            grilla.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            grilla.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grilla.StandardTab = false;
            grilla.Columns.Add("fecha", "Fecha");
            grilla.Columns.Add("comentario", "Comentario");

            grilla.Columns[0].Width = 100;
            grilla.Columns[0].ReadOnly = true;            
            grilla.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grilla.Columns[0].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
            grilla.Columns[0].DefaultCellStyle.BackColor = Color.Gray;

            grilla.Columns[1].Width = 150;
            grilla.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        }
        private void HabilitarEdicionCeldas()
        {
            //Columnas editables                
            if (RolUsuarioConectado == (int) Roles.Rol.Planif || UsuarioTieneRolAsignado())
            {
                //PlanillaOT.Columns[(int)CF.Prioridad].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.FechaEntregaRecup].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.FechaInicioEval].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.FechaFinEval].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.InfRC].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.InfEval].ReadOnly = false;
                PlanillaOT.Columns[(int)CF.InfFinal].ReadOnly = false;
                MenuDetenerTrabajo.Enabled = true;
                MenuReanudarTrabajo.Enabled = true;
            }
            else
            {
                MenuDetenerTrabajo.Enabled = false;
                MenuReanudarTrabajo.Enabled = false;
            }
        }
        private void TextoEncabezadoPlanillaOT()
        {            
            PlanillaOT.Columns[(int)CF.NumeroOT].HeaderText = K.NumeroOT;
            PlanillaOT.Columns[(int)CF.DiasEnTaller].HeaderText = K.DiasEnTaller;
            
            PlanillaOT.Columns[(int)CF.FechaRecepcion].HeaderText = K.FechaRecepcion;
            PlanillaOT.Columns[(int)CF.Avance].HeaderText = K.Avance;
            PlanillaOT.Columns[(int)CF.FechaArribo].HeaderText = K.FechaArribo;
            PlanillaOT.Columns[(int)CF.HrsEstimadas].HeaderText = K.HrsEstimadas;
            PlanillaOT.Columns[(int)CF.HrsPendientes].HeaderText = K.HrsPendientes;
            PlanillaOT.Columns[(int)CF.HrsCargadas].HeaderText = K.HrsCargadas;
            PlanillaOT.Columns[(int)CF.FechaInicioEval].HeaderText = K.FechaInicioEval;
            PlanillaOT.Columns[(int)CF.FechaFinEval].HeaderText = K.FechaFinEval;
            PlanillaOT.Columns[(int)CF.FechaFinEvalReal].HeaderText = K.FechaFinEvalReal;
            PlanillaOT.Columns[(int)CF.FechaReprogEval].HeaderText = K.FechaReprogEval;

            PlanillaOT.Columns[(int)CF.FechaInicioPlanif].HeaderText = K.FechaInicioPlanif;
            PlanillaOT.Columns[(int)CF.FechaFinPlanif].HeaderText = K.FechaFinPlanif;
            PlanillaOT.Columns[(int)CF.FechaFinPlanifReal].HeaderText = K.FechaFinPlanifReal;
            PlanillaOT.Columns[(int)CF.FechaReprogPlanif].HeaderText = K.FechaReprogPlanif;
            PlanillaOT.Columns[(int)CF.FechaInicioPpto].HeaderText = K.FechaInicioPpto;
            PlanillaOT.Columns[(int)CF.FechaFinPpto].HeaderText = K.FechaFinPpto;
            PlanillaOT.Columns[(int)CF.FechaFinPptoReal].HeaderText = K.FechaFinPptoReal;
            PlanillaOT.Columns[(int)CF.FechaReprogPpto].HeaderText = K.FechaReprogPpto;
            PlanillaOT.Columns[(int)CF.FechaInicioRep].HeaderText = K.FechaInicioRep;
            PlanillaOT.Columns[(int)CF.FechaEntregaRecup].HeaderText = K.FechaEntregaRecup;
            PlanillaOT.Columns[(int)CF.FechaInicioArmado].HeaderText = K.FechaInicioArmado;
            PlanillaOT.Columns[(int)CF.FechaFinRep].HeaderText = K.FechaFinRep;
            PlanillaOT.Columns[(int)CF.FechaFinRepReal].HeaderText = K.FechaFinRepReal;
            PlanillaOT.Columns[(int)CF.FechaEntregaProd].HeaderText = K.FechaEntregaProd;
            PlanillaOT.Columns[(int)CF.FechaLiberacion].HeaderText = K.FechaLiberacion;
            PlanillaOT.Columns[(int)CF.FechaDespacho].HeaderText = K.FechaDespacho;
            PlanillaOT.Columns[(int)CF.FechaEntregaClte].HeaderText = K.FechaEntregaClte;
            PlanillaOT.Columns[(int)CF.MotivoIncum].HeaderText = K.MotivoIncum;
            PlanillaOT.Columns[(int)CF.DiasEnProceso].HeaderText = "Dias en proceso";
            if (GetVista == Vista.Comercial)
            {
                PlanillaOT.Columns[(int)CF.FechaInicioEval].HeaderText = K.FechaInicioEvalVistas;
                PlanillaOT.Columns[(int)CF.FechaFinEval].HeaderText = K.FechaFinEvalVistas;
                PlanillaOT.Columns[(int)CF.FechaFinEvalReal].HeaderText = K.FechaFinEvalRealVistas;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].HeaderText = K.FechaFinPptoVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].HeaderText = K.FechaEntregaProdVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].HeaderText = K.FechaEntregaClteVistas;
                PlanillaOT.Columns[(int)CF.FechaFinRep].HeaderText = K.FechaFinRepVistas;            

                PlanillaOT.Columns[(int)CF.FechaFinEval].HeaderText = K.FechaFinEvalVistas;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].HeaderText = K.FechaFinPptoVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].HeaderText = K.FechaEntregaClteVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].HeaderText = K.FechaEntregaProdVistas;
                PlanillaOT.Columns[(int)CF.DiasEnProceso].HeaderText = "Dias en proceso";
                PlanillaOT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
            else if (GetVista == Vista.Comercial || GetVista == Vista.Taller || GetVista == Vista.Presupuesto || GetVista == Vista.ControlCalidad)
            {
                PlanillaOT.Columns[(int)CF.FechaInicioEval].HeaderText = K.FechaInicioEvalVistas;
                PlanillaOT.Columns[(int)CF.FechaFinEval].HeaderText = K.FechaFinEvalVistas;
                PlanillaOT.Columns[(int)CF.FechaFinEvalReal].HeaderText = K.FechaFinEvalRealVistas;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].HeaderText = K.FechaFinPptoVistas;
                PlanillaOT.Columns[(int)CF.FechaInicioPpto].HeaderText = K.FechaInicioPptoVistas;
                PlanillaOT.Columns[(int)CF.FechaFinPptoReal].HeaderText = K.FechaFinPptoRealVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].HeaderText = K.FechaEntregaProdVistas;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].HeaderText = K.FechaEntregaClteVistas;
                PlanillaOT.Columns[(int)CF.FechaFinRep].HeaderText = K.FechaFinRepVistas;
                PlanillaOT.Columns[(int)CF.DiasEnTaller].HeaderText = K.DiasEnTallerVistas;
                PlanillaOT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
               
            }                      
        }
        private void AnchoColumnasPlanillaOT()
        {

            PlanillaOT.Columns[(int)CF.NumeroOT].Width = 50;
            PlanillaOT.Columns[(int)CF.Sucursal].Width = 80;
            PlanillaOT.Columns[(int)CF.Componente].Width = 150;
            PlanillaOT.Columns[(int)CF.Cliente].Width = 150;
            PlanillaOT.Columns[(int)CF.Estado].Width = 100;
            PlanillaOT.Columns[(int)CF.Tipo].Width = 40;
            PlanillaOT.Columns[(int)CF.Prioridad].Width = 50;
            PlanillaOT.Columns[(int)CF.HrsEstimadas].Width = 60;
            PlanillaOT.Columns[(int)CF.HrsCargadas].Width = 60;
            PlanillaOT.Columns[(int)CF.HrsPendientes].Width = 60;
            PlanillaOT.Columns[(int)CF.Avance].Width = 50;            
            PlanillaOT.Columns[(int)CF.FechaRecepcion].Width = 60;
            PlanillaOT.Columns[(int)CF.DiasEnTaller].Width = 60;
            PlanillaOT.Columns[(int)CF.DiasEnProceso].Width = 60;
            PlanillaOT.Columns[(int)CF.FechaInicioEval].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinEval].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinEvalReal].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaReprogEval].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaInicioPlanif].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinPlanif].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinPlanifReal].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaReprogPlanif].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaInicioPpto].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinPpto].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinPptoReal].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaReprogPpto].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaInicioRep].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaInicioArmado].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinRep].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaFinRepReal].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaEntregaRecup].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaEntregaClte].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaEntregaProd].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaDespacho].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaLiberacion].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaDespacho].Width = 70;
            PlanillaOT.Columns[(int)CF.FechaArribo].Width = 70;
            PlanillaOT.Columns[(int)CF.Suspendida].Width = 20;
            PlanillaOT.Columns[(int)CF.CodigoFab].Width = 90;

            if (GetVista == Vista.Taller)
            {
                PlanillaOT.Columns[(int)CF.Cliente].Width = 200;
                PlanillaOT.Columns[(int)CF.FechaFinRep].Width = 100;
                PlanillaOT.Columns[(int)CF.FechaInicioEval].Width = 100;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].Width = 130;
                
            }       
            else if (GetVista == Vista.Comercial)
            {
                PlanillaOT.Columns[(int)CF.FechaFinEval].Width = 100;
                PlanillaOT.Columns[(int)CF.FechaFinPpto].Width = 100;
                PlanillaOT.Columns[(int)CF.FechaEntregaClte].Width = 120;
                PlanillaOT.Columns[(int)CF.FechaEntregaProd].Width = 130;
            }
        }
        private void FormatoFechasPlanilla()
        {
            PlanillaOT.Columns[(int)CF.FechaRecepcion].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaInicioEval].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinEval].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinEvalReal].DefaultCellStyle.Format = "dd/MM/yyyy";

            PlanillaOT.Columns[(int)CF.FechaInicioPlanif].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinPlanif].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinPlanifReal].DefaultCellStyle.Format = "dd/MM/yyyy";

            PlanillaOT.Columns[(int)CF.FechaInicioPpto].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinPpto].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinPptoReal].DefaultCellStyle.Format = "dd/MM/yyyy";

            PlanillaOT.Columns[(int)CF.FechaInicioRep].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaInicioArmado].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinRep].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaFinRepReal].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaEntregaRecup].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaArribo].DefaultCellStyle.Format = "dd/MM/yyyy";

            PlanillaOT.Columns[(int)CF.FechaEntregaClte].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaEntregaProd].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaDespacho].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaLiberacion].DefaultCellStyle.Format = "dd/MM/yyyy";
            PlanillaOT.Columns[(int)CF.FechaDespacho].DefaultCellStyle.Format = "dd/MM/yyyy";
        }
        #endregion
        private void ConfigurarGrillaHojasOT()
        {
            try
            {
                //Configuración grilla hojas de ruta
                PlanillaHojas.ReadOnly = false;
                PlanillaHojas.EditMode = DataGridViewEditMode.EditOnKeystroke;
                PlanillaHojas.AllowUserToDeleteRows = false;
                PlanillaHojas.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                PlanillaHojas.ColumnHeadersHeight = 40;                
                DataGridViewImageColumn colEstadoProc = new ()
                {
                    Name = "Img_0",
                    HeaderText = "Img"
                };

                PlanillaHojas.Columns.Add(colEstadoProc);
                PlanillaHojas.Columns.Add("HojaNum_1", "HojaNum");
                PlanillaHojas.Columns.Add("Item_2", "Item");
                PlanillaHojas.Columns.Add("Elemento_3", "Elemento");
                PlanillaHojas.Columns.Add("Seccion_4", "Seccion");
                PlanillaHojas.Columns.Add("Proceso_5", "Proceso ");
                PlanillaHojas.Columns.Add("EstadoProc_6", "EstadoProc");
                PlanillaHojas.Columns.Add("HorasEstim_7", "HorasEstim");
                PlanillaHojas.Columns.Add("Dias_8", "Dias ");
                PlanillaHojas.Columns.Add("HorasCargadas_9", "HorasCargadas");
                PlanillaHojas.Columns.Add("DiasRestantes_10", "DiasRestantes");
                PlanillaHojas.Columns.Add("FechaIni_11", "FechaIni");
                PlanillaHojas.Columns.Add("FechaFin_12", "FechaFin");
                PlanillaHojas.Columns.Add("predecesor", "Proceso Predecesor");
                PlanillaHojas.Columns.Add("CodigoEstadoProc_13", "CodigoEstadoProc");
                

                PlanillaHojas.Columns[(int)CD.HojaNum].HeaderText = "Hoja";
                PlanillaHojas.Columns[(int)CD.FechaIni].HeaderText = "Fecha Inicio";
                PlanillaHojas.Columns[(int)CD.HorasEstim].HeaderText = "Horas Estimadas";
                PlanillaHojas.Columns[(int)CD.HorasCargadas].HeaderText = "Horas Cargadas";
                PlanillaHojas.Columns[(int)CD.FechaFin].HeaderText = "Fecha Fin";
                PlanillaHojas.Columns[(int)CD.EstadoProc].HeaderText = "Estado";
                PlanillaHojas.Columns[(int)CD.Dias].HeaderText = "Dias";
                PlanillaHojas.Columns[(int)CD.DiasRestantes].HeaderText = "Dias Restantes";

                PlanillaHojas.Columns[(int)CD.HorasCargadas].DefaultCellStyle.Format = "#,##0.00";
                PlanillaHojas.Columns[(int)CD.FechaFin].DefaultCellStyle.Format = "dd/MM/yyyy";
                PlanillaHojas.Columns[(int)CD.FechaIni].DefaultCellStyle.Format = "dd/MM/yyyy";

                PlanillaHojas.Columns[(int)CD.HojaNum].Width = 40;
                PlanillaHojas.Columns[(int)CD.Predecesor].Width = 60;
                PlanillaHojas.Columns[(int)CD.Item].Width = 30;
                PlanillaHojas.Columns[(int)CD.Seccion].Width = 200;
                PlanillaHojas.Columns[(int)CD.Elemento].Width = 200;
                PlanillaHojas.Columns[(int)CD.Proceso].Width = 100;
                PlanillaHojas.Columns[(int)CD.FechaIni].Width = 70;
                PlanillaHojas.Columns[(int)CD.HorasEstim].Width = 60;
                PlanillaHojas.Columns[(int)CD.HorasCargadas].Width = 60;
                PlanillaHojas.Columns[(int)CD.FechaFin].Width = 70;
                PlanillaHojas.Columns[(int)CD.EstadoProc].Width = 100;
                PlanillaHojas.Columns[(int)CD.Dias].Width = 40;
                PlanillaHojas.Columns[(int)CD.DiasRestantes].Width = 40;
                PlanillaHojas.Columns[(int)CD.Imagen].Width = 40;
                PlanillaHojas.Columns[(int)CD.HorasEstim].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaHojas.Columns[(int)CD.Dias].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaHojas.Columns[(int)CD.DiasRestantes].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaHojas.Columns[(int)CD.HorasCargadas].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaHojas.Columns[(int)CD.Predecesor].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                PlanillaHojas.Columns[(int)CD.CodigoEstadoProc].Visible = false;                
                
                //Columnas solo lectura
                for (int i = 0; i < PlanillaHojas.ColumnCount; i++)
                {
                    PlanillaHojas.Columns[i].ReadOnly = true;
                    PlanillaHojas.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                PlanillaHojas.Columns[(int)CD.Predecesor].ReadOnly = false;
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void ConfigurarGrillaTrabajos()
        {
            try
            {                
                GrillaTrabajos.ReadOnly = true;
                GrillaTrabajos.AllowUserToResizeColumns = true;
                
                GrillaTrabajos.Columns[(int)ColTrab.FechaIni].DefaultCellStyle.Format = "dd/MM/yyyy";
                GrillaTrabajos.Columns[(int)ColTrab.FechaFin].DefaultCellStyle.Format = "dd/MM/yyyy";
                GrillaTrabajos.Columns[(int)ColTrab.FechaIni].HeaderText = "Fecha Inicio";
                GrillaTrabajos.Columns[(int)ColTrab.FechaFin].HeaderText = "Fecha Fin";
                GrillaTrabajos.Columns[(int)ColTrab.HoraIni].HeaderText = "Hora Inicio";
                GrillaTrabajos.Columns[(int)ColTrab.HoraFin].HeaderText = "Hora Fin";
                GrillaTrabajos.Columns[(int)ColTrab.TiempoTrans].HeaderText = "Cumplido";
                GrillaTrabajos.Columns[(int)ColTrab.TiempoEstim].HeaderText = "Estimado";

                GrillaTrabajos.Columns[(int)ColTrab.Trabajador].Width = 200;
                GrillaTrabajos.Columns[(int)ColTrab.NumOT].Width = 60;
                GrillaTrabajos.Columns[(int)ColTrab.NumHoja].Width = 40;
                GrillaTrabajos.Columns[(int)ColTrab.Proceso].Width = 150;
                GrillaTrabajos.Columns[(int)ColTrab.FechaIni].Width = 80;
                GrillaTrabajos.Columns[(int)ColTrab.FechaFin].Width = 80;
                GrillaTrabajos.Columns[(int)ColTrab.HoraIni].Width = 65;
                GrillaTrabajos.Columns[(int)ColTrab.HoraFin].Width = 60;
                GrillaTrabajos.Columns[(int)ColTrab.TiempoTrans].Width = 60;
                GrillaTrabajos.Columns[(int)ColTrab.TiempoEstim].Width = 70;              
                GrillaTrabajos.Columns[(int)ColTrab.HoraIni].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                GrillaTrabajos.Columns[(int)ColTrab.HoraFin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                GrillaTrabajos.Columns[(int)ColTrab.TiempoTrans].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                GrillaTrabajos.Columns[(int)ColTrab.TiempoEstim].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                GrillaTrabajos.Columns[(int)ColTrab.NumOT].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                GrillaTrabajos.Columns[(int)ColTrab.NumHoja].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //GrillaTrabajos.Columns[(int)CD.CodigoEstadoProc].Visible = false;


            }
            catch (Exception ex)
            {
                SetErrorMsg(ex.Message);
            }
        }
        private void InicializarValoresControles()
        {
            //Inicializa controles de fecha con fecha de hoy
            FechaEntregaIni.Value = DateTime.Now;
            FechaEntregaFin.Value = DateTime.Now;
            FechaEntregaProduccionIni.Value = DateTime.Now;
            FechaEntregaProduccionFin.Value = DateTime.Now;            
            MenuAbrirHojaRuta.Enabled = false;            
            timer1.Interval = INTERVALO_ACTUALIZACION;
            
            EstadosTaller.Checked = true;
            Font F1 = new(PanelMsg.Font.Name, PanelMsg.Font.Size, FontStyle.Regular);
            Font F2 = new(PanelMsg.Font.Name, 12, FontStyle.Bold);
            PanelMsg.Font = F2;
            Funciones.Mensajes(PanelMsg, MsgCarga, esError: false);

            PanelMsg.Font = F1;
            Sucursal.Enabled = true;
            Seccion.Enabled = true;
            PanelMsg.Text = string.Empty;
            TipoSec = Capacidades.TipoSeccion.Armado;         
            tsslVersion.Text = Funciones.VersionAssembly();

            LblSuspend.BackColor = ColorSuspendida;
            LblAtraso.BackColor = ColorAporteClte;
            LblFuera.BackColor = ColorFueraPlazo;
            LblRetraso.BackColor = ColorRetrasado;
            LblHH.BackColor = ColorCargaHHMayorEstima;            

            ConfiguracionConexion.LeerAppConfig(BaseDatos.Parametro(BaseDatos.Parametros.NombreModeloEF));
            PanelConex.Text = "Base de datos: " + ConfiguracionConexion.BaseDeDatos + " en " + ConfiguracionConexion.Servidor;
            if (GetVista != Vista.Armado && GetVista != Vista.Recuperacion)
            {
                FicheroPlanificacion.TabPages.RemoveAt(1);
                FicheroPlanificacion.TabPages.RemoveAt(1);

            }
            if (GetVista == Vista.Comercial)
            {
                splitContainer1.Panel2Collapsed = true;
                MenuVentas.Checked = true;
                FicheroPlanificacion.TabPages.RemoveAt(1);
            }           
            else if (GetVista == Vista.Armado)
            {
                TituloEjeX = "Secciones desarme y armado";
                TituloGrafico = "Capacidad instalada secciones desarme y armado";
            }
            else if (GetVista == Vista.Recuperacion)
            {
                TituloEjeX = "Secciones fabricación y recuperación";
                TituloGrafico = "Capacidad instalada secciones fabricación y recuperación";
            }
        }
        private bool ListarDatos()
        {
            Cursor = Cursors.WaitCursor;
            chkVendedor.Enabled = true;
            PlanillaHojas.DataSource = null;
            if (ListarPlanificacion())
            {
                ObtenerCapacidadesOT();
                CumplimientoPlazos();
                tsbImprimir.Enabled = true;
                Cursor = Cursors.Default;
                return true;
            }            
            else
            {
                Cursor = Cursors.Default;
                return false;
            }
                
            
        }       
        private List<int> DiasProceso(int CodigoTipo,int CodigoEstado,DateTime? FechaRecepcionComponente, DateTime? FechaFinPptoReal,DateTime? FechaInicioRep)
        {
            var pl = new Planificacion.Planificacion();
            var lsta = pl.DiasComponenteEnTaller(CodigoTipo, CodigoEstado, FechaRecepcionComponente, FechaFinPptoReal, FechaInicioRep);
            if (lsta != null && lsta.Count > 0)
            {
                return lsta;
            }
            return null;
        }
        private bool ListarPlanificacion()
        {
            try
            {
                string msg = string.Empty;

                var pl = new Planificacion.Planificacion
                {
                    NumeroOT1 = NumeroOTIni.Text == string.Empty ? 0 : Convert.ToInt32(NumeroOTIni.Text),
                    NumeroOT2 = NumeroOTFin.Text == string.Empty ? 0 : Convert.ToInt32(NumeroOTFin.Text),
                    Nivel1 = Convert.ToInt16(Sucursal.SelectedValue),
                    Nivel2 = Convert.ToInt16(Negocio.SelectedValue),
                    Nivel3 = Convert.ToInt16(Producto.SelectedValue),
                    Nivel5 = Convert.ToInt16(Seccion.SelectedValue),
                    FechaEntregaSeleccionada = FechaEntregaClienteCheck.Checked == true ? 1 : 0,
                    FechaDesde = FechaEntregaClienteCheck.Checked ? FechaEntregaIni.Value : DateTime.MinValue,
                    FechaHasta = FechaEntregaFin.Value,
                    FechaEntregaProdSeleccionada = FechaEntregaProduccionCheck.Checked == true ? 1 : 0,
                    FechaDesdeProd = FechaEntregaProduccionCheck.Checked ? FechaEntregaProduccionIni.Value : DateTime.MinValue,
                    FechaHastaProd = FechaEntregaProduccionFin.Value,
                    ResponsableOT = Convert.ToInt32(ResponsableOT.SelectedValue),
                    ListaEstadosOT = EstadosOTSel(),
                    ListaTiposOT = TiposOTSel(),
                    MostrarOTActiva = !MostrarOtActivas.Checked,
                    MostrarSoloOTDetenida = !MostrarOtDetenidas.Checked,                    
                    IncluirOtEvalEnFechaEntrega = MenuFecEntrEval.Checked,
                    RutCliente = Cliente.SelectedValue != null ? Cliente.SelectedValue.ToString() : string.Empty,
                    CodigoSucCliente = SucursalCliente.SelectedValue != null ? int.Parse(SucursalCliente.SelectedValue.ToString()) : 0,
                    CodigoEstadoEntrega = EstadoEntrega.SelectedValue != null ? byte.Parse(EstadoEntrega.SelectedValue.ToString()) : (byte)0,
                };
                
                var lst = pl.ListarOTsPlanificacion(ref msg);                
                if (lst != null)
                {
                    TablaOTs = lst;
                    LlenarGrillaPlanificacion(lst);
                    return true;                                              
                }
                else
                {                                        
                    SetErrorMsg(msg);
                    return false;
                }
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
                return false;
            }
        }
        private void ObtenerSeguimientoProcesos(int fila, DataGridView grilla, TipoSeguimiento.TiposDeSeguimiento tipo)
        {
            grilla.Rows.Clear();            
            string msg = string.Empty;
            int.TryParse(PlanillaOT[(int)CF.NumeroOT, fila].Value.ToString(), out int numero);
            var pl = new Planificacion.Planificacion()
            {
                NumeroOT1 = numero
            };
            var data = pl.ComentariosSeguimiento(tipo, ref msg);
            foreach (var item in data)
            {
                grilla.Rows.Add(item.Fecha, item.Comentario);
                grilla.Rows[grilla.RowCount - 1].DefaultCellStyle.BackColor = Color.Gray;
                grilla.Rows[grilla.RowCount - 1].ReadOnly = true;
            }
            grilla.RowCount += 1;

        }
        private void ObtenerDatosVentaSAP()
        {            
            double totalVenta = 0;
            string ruta = string.Empty;
            string sesion = string.Empty;
            var pl = new Planificacion.Planificacion();
            int i = 0;
            BarraProg.Visible = true;
            PanelMsg.Text = "Obteniendo datos orden de venta...";
            BarraProg.Maximum = TablaOTs.Count;
            foreach (var item in TablaOTs)
            {
                Application.DoEvents();
                var lst = pl.ListarDatosOVentaSap(item.Numero, ref sesion, ref ruta);
                if (lst != null && lst.Count > 0)
                {
                    item.Vendedor = lst[0];
                    item.OCcliente = lst[1];
                    double.TryParse(lst[2], out double valor);                   

                    PlanillaOT[(int)CF.Vendedor, i].Value = item.Vendedor;
                    PlanillaOT[(int)CF.OcCliente, i].Value = item.OCcliente;
                    PlanillaOT[(int)CF.ValorTrabajo, i].Value = string.Format("{0:#,##0}", valor / IMPUESTO);
                }
                else
                {
                    item.Vendedor = pl.VendedorActividad(item.Numero, sesion, ruta);
                    string oferta = pl.OfertaVenta(item.Numero, sesion, ruta);                    

                    PlanillaOT[(int)CF.Vendedor,i].Value = item.Vendedor;
                    PlanillaOT[(int)CF.OfertaVenta, i].Value = oferta;
                }
                totalVenta += pl.MontoVenta;
                i += 1;
                BarraProg.Value = i;
            }            
            TotalVentaPeriodo.Text = string.Format("{0:#,##0}", totalVenta / IMPUESTO);
            BarraProg.Visible = false;
            PanelMsg.Text = string.Empty;
        }                
        private void AgregarVendedorCombo()
        {
            var dt = new DataTable();
            dt.Columns.Add("Codigo");
            dt.Columns.Add("Descripcion");
            foreach (DataGridViewRow dr in PlanillaOT.Rows)
            {
                dt.Rows.Add(dr.Cells[(int)CF.Vendedor].Value, dr.Cells[(int)CF.Vendedor].Value);
            }
            var newDt = dt.DefaultView.ToTable(true, dt.Columns[0].ColumnName, dt.Columns[1].ColumnName);
            var dv = newDt.DefaultView;
            dv.Sort = dt.Columns[0].ColumnName;
            Vendedor.DataSource = dv;
            Vendedor.DisplayMember = "Descripcion";
            Vendedor.ValueMember = "Codigo";
            Vendedor.SelectedIndex = -1;
            chkVendedor.Enabled = true;
        }
        private void LlenarGrillaPlanificacion(SortableBindingList<GestionOrdenTrabajo.AuxGestion> listado)
        {
            var pl = new Planificacion.Planificacion();
            PlanillaOT.RowCount = 0;
            int diasTallerEval, diasTallerRepar;
            foreach (var item in listado)
            {
                diasTallerEval = 0;
                diasTallerRepar = 0;
                var lsta = DiasProceso(item.CodigoTipo, item.CodigoEstado, item.FechaRecepcionComponente, item.FechaFinPptoReal, item.FechaInicioRep);
                if (lsta != null)
                {
                    diasTallerEval = lsta[0];
                    diasTallerRepar = lsta[1];
                }
                string mip = $"{item.MipPpto}{item.MipEval}{item.MipEntr}{item.MipPlan}{item.MipRepa}";
                int avance = pl.AvanceTrabajo(item.Numero, item.CodigoEstado);
                PlanillaOT.Rows.Add(item.Numero, item.Nivel1, item.Componente, item.Marca, item.Modelo, item.Cliente, item.TipoCorto, item.Estado,
                    item.Prioridad, avance, 0, 0, 0, item.FechaRecepcionComponente, diasTallerEval, diasTallerRepar,
                    item.FechaInicioEval, item.FechaFinEval, item.FechaFinEvalReal, item.FechaRepogramacionEval,
                    item.FechaInicioPlanif, item.FechaFinPlanif, item.FechaFinPlanifReal,item.FechaRepogramacionPlanif,
                    item.FechaInicioPpto, item.FechaFinPpto, item.FechaFinPptoReal, item.FechaRepogramacionPpto,
                    item.FechaInicioRep, item.FechaEntregaRecup, item.FechaInicioArmado, item.FechaFinRep, item.FechaFinRepReal, item.FechaArribo,
                    item.FechaEntregaPlanificacion, item.FechaLiberacion, item.FechaDespacho, item.FechaEntregaCliente,
                    item.OTSuspendida, item.CodigoEstado, item.Detalle, item.Responsable, item.Asistente,
                    item.NombreDossier, mip, item.InfRC, item.InfEV, item.InfFin,
                    item.Vendedor, item.OCcliente, item.ValorTrabajo, item.OfertaVenta, item.CodigoEstadoEntrega,
                    item.CodigoArticuloFabricado, item.AporteCliente);
            }
            tsslRegistros.Text = listado.Count.ToString() + " órdenes de trabajo";
            
        }
        private void FiltraPorVendedor()
        {
            try
            {
                var pl = new Planificacion.Planificacion
                {
                    Usuario = UsuarioConectado
                };
                PlanillaOT.RowCount = 0;
                if (Vendedor.Text != string.Empty)
                {
                    var lst = (from dt in TablaOTs
                              where dt.Vendedor == Vendedor.Text
                              select dt).ToList();
                    SortableBindingList<GestionOrdenTrabajo.AuxGestion> lista = new(lst);
                    LlenarGrillaPlanificacion(lista);
                }
                else
                {                    
                    var lst = (from dt in TablaOTs                               
                               select dt).ToList();
                    SortableBindingList<GestionOrdenTrabajo.AuxGestion> lista = new(lst);
                    LlenarGrillaPlanificacion(lista);                   
                }                               
                ObtenerCapacidadesOT();                
                CumplimientoPlazos();
            }
            catch (Exception ex)
            {                                
                SetErrorMsg(ex.Message);
            }
        }
        private void ObtenerCapacidadesOT()
        {
            List<int> listaOT = new();
            
            foreach (var item in TablaOTs)
            {
                listaOT.Add(item.Numero);
            }
            var cap = new Capacidades();
            var listaCapacidades = cap.CapacidadPorOT(listaOT, TipoSec);

            foreach (DataGridViewRow dr in PlanillaOT.Rows)
            {
                int numOT = int.Parse(dr.Cells[(int)CF.NumeroOT].Value.ToString());
                var qry = listaCapacidades.Find(x => x.NumeroOT == numOT);
                if (qry != null)
                {
                    dr.Cells[(int)CF.HrsEstimadas].Value = qry.HorasEstimadasHorario;
                    dr.Cells[(int)CF.HrsCargadas].Value =  qry.HorasCargadasHorario;
                    dr.Cells[(int)CF.HrsPendientes].Value = qry.HorasPorCargarHorario;
                }                          
            }            
        }        
        private bool UsuarioTieneRolAsignado()
        {
            var plan = new Planificacion.Planificacion
            {
                Usuario = UsuarioConectado
            };
            bool usuarioTieneRol = plan.RolAsignado(RolUsuarioConectado);
            return usuarioTieneRol;
        }
        private void ResaltaOTDetenida()
        {
            for (int i = 0; i < PlanillaOT.RowCount; i++)
            {
                bool detenida = bool.Parse(PlanillaOT[(int)CF.Suspendida, i].Value.ToString());
                if (detenida)
                    PlanillaOT.Rows[i].DefaultCellStyle.BackColor = Color.CadetBlue;
            }
        }
        protected void CumplimientoPlazos()
        {                                                          
            for (int i = 0; i < PlanillaOT.RowCount; i++)
            {
                int codigoEstado = Convert.ToInt16(PlanillaOT[(int)CF.CodigoEstado, i].Value);
                // Color rojo pendientes < 0 y OT's suspendidas                
                bool otSuspendida = Convert.ToBoolean(PlanillaOT[(int)CF.Suspendida, i].Value);
                bool otAporteClte = Convert.ToBoolean(PlanillaOT[(int)CF.AporteClte, i].Value);
                _ = byte.TryParse(PlanillaOT[(int)CF.EstadoEntrega, i].Value.ToString(),out byte codEntrega);
                string horasDec = PlanillaOT[(int)CF.HrsPendientes, i].Value.ToString();
                double pdtes = 0;
                if (horasDec != "0")
                    pdtes = Funciones.ConvertirHoraSexagecimalEnDecimal(horasDec);
                if (pdtes < 0)
                {
                    //Celda color
                    PlanillaOT.Rows[i].Cells[(int)CF.HrsPendientes].Style.BackColor = ColorCargaHHMayorEstima;                    
                }
                if (otSuspendida == true)
                    //Fila color
                    PlanillaOT.Rows[i].DefaultCellStyle.BackColor = ColorSuspendida;                
                if (codEntrega == 2)        //Retrasado        
                {
                    //Fila color
                    PlanillaOT.Rows[i].DefaultCellStyle.BackColor = ColorRetrasado;                    
                    
                }
                else if (codEntrega == 3) //Fuera de plazo
                    //Fila color
                    PlanillaOT.Rows[i].DefaultCellStyle.BackColor = ColorFueraPlazo;
                if (otAporteClte == true)
                    //Fila color
                    PlanillaOT.Rows[i].DefaultCellStyle.BackColor = ColorAporteClte;
                //Celda color
                if (codigoEstado == (int) EstadOrdenTrabajo.Estados.Evaluacion || codigoEstado == (int)EstadOrdenTrabajo.Estados.Proceso 
                    || codigoEstado == (int)EstadOrdenTrabajo.Estados.Planificacion)
                {
                    if (PlanillaOT[(int)CF.FechaFinEval, i].Value != DBNull.Value)
                        CumplimientoPlazosEvaluacion(i);
                    if (PlanillaOT[(int)CF.FechaFinPlanif, i].Value != DBNull.Value)
                        CumplimientoPlazosPlanificacion(i);
                    if (PlanillaOT[(int)CF.FechaFinPpto, i].Value != DBNull.Value)
                        CumplimientoPlazosPresupuesto(i);
                    if (PlanillaOT[(int)CF.FechaFinRep, i].Value != DBNull.Value)
                        CumplimientoPlazosReparacion(i);
                }
                if (PlanillaOT[(int)CF.FechaEntregaClte, i].Value != DBNull.Value)
                {
                    CumplimientoPlazosEntrega(i);
                }
            }
        }        
        private void CumplimientoPlazosEvaluacion(int i)
        {
            TimeSpan ts;
            DateTime fechaTerminoReal;

            DateTime fechaTerminoEval = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinEval, i].Value);
            if (PlanillaOT[(int)CF.FechaFinEvalReal, i].Value != DBNull.Value)
            {
                fechaTerminoReal = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinEvalReal, i].Value);
                ts = fechaTerminoReal.Subtract(fechaTerminoEval);
                if (ts.Ticks > 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinEval].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinEval].Style.ForeColor = ColorLetra;
                }
            }
            else
            {
                ts = fechaTerminoEval.Subtract(DateTime.Now.Date);
                if (ts.Ticks < 0) 
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinEval].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinEval].Style.ForeColor = ColorLetra;
                }                                
            }
        }
        private void CumplimientoPlazosPlanificacion(int i)
        {
            TimeSpan ts;
            DateTime fechaTerminoReal;

            DateTime fechaTerminoPlanif = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinPlanif, i].Value);
            if (PlanillaOT[(int)CF.FechaFinPlanifReal, i].Value != DBNull.Value)
            {
                fechaTerminoReal = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinPlanifReal, i].Value);
                ts = fechaTerminoReal.Subtract(fechaTerminoPlanif);
                if (ts.Ticks > 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPlanif].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPlanif].Style.ForeColor = ColorLetra;
                }
            }
            else           
            {
                ts = fechaTerminoPlanif.Subtract(DateTime.Now.Date);
                if (ts.Ticks < 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPlanif].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPlanif].Style.ForeColor = ColorLetra;
                }
            }            
        }
        private void CumplimientoPlazosPresupuesto(int i)
        {
            TimeSpan ts;
            DateTime fechaTerminoReal;

            DateTime fechaTerminoPpto = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinPpto, i].Value);
            if (PlanillaOT[(int)CF.FechaFinPptoReal, i].Value != DBNull.Value)
            {
                fechaTerminoReal = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinPptoReal, i].Value);
                ts = fechaTerminoReal.Subtract(fechaTerminoPpto);
                if (ts.Ticks > 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPpto].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPpto].Style.ForeColor = ColorLetra;
                }
            }
            else
            {
                ts = fechaTerminoPpto.Subtract(DateTime.Now.Date);
                if (ts.Ticks < 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPpto].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinPpto].Style.ForeColor = ColorLetra;
                }
            }
           
        }
        private void CumplimientoPlazosReparacion(int i)
        {
            TimeSpan ts;
            DateTime fechaTerminoReal;

            DateTime fechaTerminoRepar = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinRep, i].Value);
            if (PlanillaOT[(int)CF.FechaFinRepReal, i].Value != DBNull.Value)
            {
                fechaTerminoReal = Convert.ToDateTime(PlanillaOT[(int)CF.FechaFinRepReal, i].Value);
                ts = fechaTerminoReal.Subtract(fechaTerminoRepar);
                if (ts.Ticks > 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinRep].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinRep].Style.ForeColor = ColorLetra;
                }
            }
            else            
            {
                ts = fechaTerminoRepar.Subtract(DateTime.Now.Date);
                if (ts.Ticks < 0)
                {
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinRep].Style.BackColor = ColorAtraso;
                    PlanillaOT.Rows[i].Cells[(int)CF.FechaFinRep].Style.ForeColor = ColorLetra;
                }
            }           
        }
        private void CumplimientoPlazosEntrega(int i)
        {
            TimeSpan ts;           
            DateTime fechaEntregaClte = Convert.ToDateTime(PlanillaOT[(int)CF.FechaEntregaClte, i].Value);
            DateTime fechaEntregaProd = Convert.ToDateTime(PlanillaOT[(int)CF.FechaEntregaProd, i].Value);

            if (fechaEntregaProd == DateTime.MinValue && fechaEntregaClte == DateTime.MinValue)
                return;
            if (fechaEntregaProd != DateTime.MinValue)
            {                
                if (fechaEntregaClte != DateTime.MinValue)
                    ts = fechaEntregaProd.Subtract(fechaEntregaClte);
                else                    
                    ts = DateTime.Now.Subtract(fechaEntregaProd);
            }
            else
                ts = DateTime.Now.Subtract(fechaEntregaClte);

            if (ts.Ticks > 0)
            {
                PlanillaOT.Rows[i].Cells[(int)CF.FechaEntregaClte].Style.BackColor = ColorAtraso;
                PlanillaOT.Rows[i].Cells[(int)CF.FechaEntregaClte].Style.ForeColor = ColorLetra;
            }
        }
        private bool ListarHojasOT(int fila)
        {
            //Lista hojas de ruta de OT seleccionada en lista de OT's
            try
            {
                PanelMsg.Text = string.Empty;
                string msg = string.Empty;
                int est = 0, numOT = 0;
                List<string> estadosPP = new() ;

                _ = EstadoProceso.SelectedValue != null ? int.TryParse(EstadoProceso.SelectedValue.ToString(), out est) ? 0 : est : 0;
                _ = PlanillaOT[(int)CF.NumeroOT, fila].Value != null ? int.TryParse(PlanillaOT[(int)CF.NumeroOT, fila].Value.ToString(), out numOT) ? 0 : numOT : 0;
                if (numOT == 0) return false;

                var pl = new Planificacion.Planificacion
                {
                    NumeroOT1 = numOT,
                    EstadoProceso = est,
                    TipoSeccion = TipoSec,
                    MostrarOTActiva = true,
                    MostrarSoloOTDetenida = false
                };
                if (est > 0)
                    estadosPP.Add(est.ToString());
                var dt = pl.ListarHojasDeRuta(ref msg, estadosPP);
                if (dt != null)
                {
                    PlanillaHojas.RowCount = 0;                    
                    
                    foreach (var item in dt)
                    {
                        decimal hhCargadas = ObtenerHorasCargadasProcesos(item.NumeroOT, item.Numero, item.Item);
                        PlanillaHojas.Rows.Add(null, item.Numero, item.Item, item.Elemento, item.Seccion, item.Proceso, item.EstadoProceso,
                            item.HHEstim, Math.Ceiling((decimal)(item.HHEstim/8)), hhCargadas, 0,
                            item.FechaInicio, item.FechaFin,item.ProcesoPredecesor, item.CodigoEstadoProceso);
                        
                    }                   
                    AlternarColorProcesoHoja();
                    if (PlanillaHojas.RowCount > 0)
                    {             
                        DestacarHestimadasMayorHcargadas();
                    }                    
                    return true;
                }
                else
                {                                        
                    SetErrorMsg(msg);
                    return false;
                }
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
                return false;
            }
        }
        private decimal ObtenerHorasCargadasProcesos(int numOT, int numeroHoja, int itemHoja)
        {
            var hh = new HojaRuta.HojaRutaDetalle(numOT)
            {                
                Numero = numeroHoja,
                Item = itemHoja
            };
            decimal horas = hh.HorasCargadasProceso();
            return horas;
        }
        private void AlternarColorProcesoHoja()
        {
            Color color1 = Color.LightGreen;
            Color color2 = Color.Ivory;
            Color color = color1;
            int hoja_sgte = 0;
            int hoja_anterior = 0;
            int predecesor = 0;
            byte cambioColor = 0;
            for (int i = 0; i < PlanillaHojas.RowCount; i++)
            {
                int hoja = int.Parse(PlanillaHojas[(int)CD.HojaNum, i].Value.ToString());

                if (i < PlanillaHojas.RowCount - 1)
                    hoja_sgte = int.Parse(PlanillaHojas[(int)CD.HojaNum, i + 1].Value.ToString());
                if (i > 0)
                    hoja_anterior = int.Parse(PlanillaHojas[(int)CD.HojaNum, i - 1].Value.ToString());               
                predecesor += 1;
                //Cambio
                if (hoja != hoja_anterior && hoja == hoja_sgte || hoja != hoja_anterior && hoja != hoja_sgte)
                {
                    predecesor = 0;
                    if (cambioColor == 0)
                    {
                        color = color2;
                        cambioColor = 1;
                    }
                    else
                    {
                        color = color1;
                        cambioColor = 0;
                    }
                }
                PlanillaHojas.Rows[i].DefaultCellStyle.BackColor = color;
            }
        }        
        private Dictionary<int, int> ObtenerProcesosParalelos(int hoja)
        {            
            Dictionary<int, int> items = new();            
            for (int i = 0; i < PlanillaHojas.RowCount; i++)
            {
                int hoja_sgte = int.Parse(PlanillaHojas[(int)CD.HojaNum, i].Value.ToString());
                if (hoja_sgte == hoja)
                {
                    int item = int.Parse(PlanillaHojas[(int)CD.Item, i].Value.ToString());
                    int numeroProc = int.Parse(PlanillaHojas[(int)CD.Predecesor, i].Value.ToString());
                    items.Add(item, numeroProc);
                    if (i < PlanillaHojas.RowCount - 1)
                    {
                        hoja_sgte = int.Parse(PlanillaHojas[(int)CD.HojaNum, i + 1].Value.ToString());
                    }
                    //Cambio
                    if (hoja != hoja_sgte)
                    {
                        break;
                    }
                }                
            }
            return items;
        }
        private void DestacarHestimadasMayorHcargadas()
        {
            //Marcar en rojo cuando HH estimadas < HH cargadas
            for (int i = 0; i < PlanillaHojas.RowCount; i++)
            {
                double HorasCargadas = 0;
                double HorasEstimadas = double.Parse(PlanillaHojas[(int)CD.HorasEstim, i].Value.ToString());
                if (PlanillaHojas[(int)CD.HorasCargadas, i].Value != DBNull.Value)
                    HorasCargadas = double.Parse(PlanillaHojas[(int)CD.HorasCargadas, i].Value.ToString());

                if (HorasEstimadas < HorasCargadas)
                {
                    PlanillaHojas.Rows[i].Cells[(int)CD.HorasCargadas].Style.BackColor = ColorCargaHHMayorEstima;
                }
            }
        }
        private void ColorProceso(DataGridView dg)
        {
            //Color indicador de estado proceso, con imagen en columna 1 del lsita de hojas de ruta
            try
            {               
                //Alternar color para hojas de ruta               
                Color color1 = Color.LightGreen;
                Color color2 = Color.Ivory;
                Color color = color1;                
                byte codigoProc = 0;

                for (int i = 0; i < dg.RowCount; i++)
                {
                    codigoProc = byte.Parse(dg[(int)CD.CodigoEstadoProc, i].Value.ToString());
                    switch (codigoProc)
                    {
                        case 1:
                            dg[(int)CD.Imagen, i].Value = Properties.Resources.redled;
                            break;
                        case 2:
                            dg[(int)CD.Imagen, i].Value = Properties.Resources.yellowled;
                            break;
                        case 3:
                            dg[(int)CD.Imagen, i].Value = Properties.Resources.greenled;
                            break;
                        default:
                            break;
                    }                                        
                }
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void HabilitarControles(object sender, bool estado)
        {
            Type t = sender.GetType();
            if (t.Equals(typeof(TextBox)))
            {
                TextBox txt;
                txt = (TextBox)sender;
                txt.Enabled = estado;
                txt.Text = string.Empty;
                txt.Focus();
            }
            if (t.Equals(typeof(Button)))
            {
                Button txt;
                txt = (Button)sender;
                txt.Enabled = estado;
            }
            if (t.Equals(typeof(ComboBox)))
            {
                ComboBox txt;
                txt = (ComboBox)sender;
                txt.Enabled = estado;                
                txt.SelectedIndex = -1;                
                txt.Focus();
            }
            if (t.Equals(typeof(DateTimePicker)))
            {
                DateTimePicker txt;
                txt = (DateTimePicker)sender;
                txt.Enabled = estado;
                txt.Value = DateTime.Now;
                txt.Focus();
            }
            if (t.Equals(typeof(CheckedListBox)))
            {
                CheckedListBox txt;
                txt = (CheckedListBox)sender;
                txt.Enabled = estado;               
                txt.Focus();
                foreach (int i in txt.CheckedIndices)
                {
                    txt.SetItemCheckState(i, CheckState.Unchecked);
                }

            }
        }
        
        #region Actualizaciones
        private void ActualizarCampos(int fila)
        {
            //Actualiza campos prioridad, fecha entrega produccion,fecha entrega recuperación y fechas inicio y término evaluación
            //Fuera de uso por acuerdo entre participantes
            try
            {
                var pl = new ActualizacionesPlanificacion();                
                bool retorno = true;
                bool fechaCorrecta = true;
                string msg = string.Empty;

                int col = 0;
                DateTime fecha;

                col = PlanillaOT.CurrentCell.ColumnIndex;
                pl.NumeroOT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, fila].Value);
                switch (col)
                {                   
                    //Actualizar prioridad
                    //case (int)CF.Prioridad:
                    //    pl.Prioridad = byte.Parse(PlanillaOT[(int)CF.Prioridad, fila].Value.ToString());                        
                    //    retorno = pl.ActualizarPrioridad();
                    //    break;
                    case (int)CF.FechaEntregaProd:
                        fechaCorrecta = DateTime.TryParse(PlanillaOT[(int)CF.FechaEntregaProd, fila].Value.ToString(), out fecha);
                        if (fechaCorrecta)
                        {
                            pl.FechaEntregaProduccion = fecha;
                            int comparaFecha = fecha.CompareTo(DateTime.Today);
                            if (comparaFecha < 0)                                
                                SetErrorMsg(MsgErrorFechaProd);
                            else
                            {
                                retorno = pl.ActualizaFechaEntregaProduccion();
                                if (retorno)
                                    PanelMsg.Text = MsgOK;
                            }
                        }
                        break;
                    case (int)CF.FechaEntregaRecup:
                        fechaCorrecta = DateTime.TryParse(PlanillaOT[(int)CF.FechaEntregaRecup, fila].Value.ToString(), out fecha);
                        if (fechaCorrecta)
                        {
                            pl.FechaEntregaRecuperacion = fecha;
                            int comparaFecha = fecha.CompareTo(DateTime.Today);
                            if (comparaFecha < 0)                                
                                SetErrorMsg(MsgErrorFechaRecup);
                            else
                            {
                                retorno = pl.ActualizaFechaEntregaRecuperacion();
                                if (retorno)
                                {
                                    PanelMsg.Visible = true;
                                    PanelMsg.Text = MsgOKRecup;
                                }
                            }
                        }
                        break;
                    case (int)CF.FechaInicioEval:
                        fechaCorrecta = DateTime.TryParse(PlanillaOT[(int)CF.FechaInicioEval, fila].Value.ToString(), out fecha);
                        if (fechaCorrecta)
                        {
                            pl.FechaInicioEvaluacion = fecha;
                            retorno = pl.ActualizarFechaInicioEvaluacion();
                            if (retorno)
                            {
                                PanelMsg.Visible = true;
                                PanelMsg.Text = MsgFechaEvalIni;
                            }
                        }
                        break;
                    case (int)CF.FechaFinEval:
                        fechaCorrecta = DateTime.TryParse(PlanillaOT[(int)CF.FechaFinEval, fila].Value.ToString(), out fecha);
                        if (fechaCorrecta)
                        {
                            pl.FechaTerminoEvaluacion = fecha;
                            PanelMsg.Visible = true;
                            PanelMsg.Text = MsgFechaEvalFin;
                            retorno = pl.ActualizarFechaFinEvaluacion();
                        }
                        break;
                    default:
                        break;
                }
                timer1.Enabled = true;
                if (retorno == false)                    
                    SetErrorMsg(msg);
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void ActualizarEstadoInformes()
        {
            PlanillaOT.CommitEdit(DataGridViewDataErrorContexts.Commit);
            int fila = PlanillaOT.CurrentRow.Index;
            int columnaOT = PlanillaOT.CurrentCell.ColumnIndex;
            int.TryParse(PlanillaOT[(int)CF.NumeroOT, fila].Value.ToString(), out int numOT);
            var plan = new ActualizacionesPlanificacion();
            bool estado;

            if (columnaOT == (int)CF.InfRC)
            {
                estado = PlanillaOT[(int)CF.InfRC, fila].Value != null && bool.Parse(PlanillaOT[(int)CF.InfRC, fila].Value.ToString());
                plan.ActualizarEstadoEntregaInformesTecnicos(numOT, estado, OrdenTrabajo.OrdenTrabajo.TipoInformeTecnico.InfRecepcion);
            }
            else if (columnaOT == (int)CF.InfEval)
            {
                estado = PlanillaOT[(int)CF.InfEval, fila].Value != null && bool.Parse(PlanillaOT[(int)CF.InfEval, fila].Value.ToString());
                plan.ActualizarEstadoEntregaInformesTecnicos(numOT, estado, OrdenTrabajo.OrdenTrabajo.TipoInformeTecnico.InfEvaluacion);
            }
            else if (columnaOT == (int)CF.InfFinal)
            {
                estado = PlanillaOT[(int)CF.InfFinal, fila].Value != null && bool.Parse(PlanillaOT[(int)CF.InfFinal, fila].Value.ToString());
                plan.ActualizarEstadoEntregaInformesTecnicos(numOT, estado, OrdenTrabajo.OrdenTrabajo.TipoInformeTecnico.InfFinal);
            }                        
        }
        private void ActualizarHHEstimadas()
        {
            //Actualiza horas estimadas del proceso seleccionado            

            int fila = PlanillaHojas.CurrentRow.Index;
            int estadoProceso = Convert.ToInt16(PlanillaHojas[(int)CD.CodigoEstadoProc, fila].Value);
            if (estadoProceso != (int)EstadoProcesoProduccion.Estados.Ejecucion)
            {                
                SetErrorMsg(MsgAviso);
                return;
            }
            else
            {
                int FilaOT = PlanillaOT.CurrentCell.RowIndex;
                int Numero_OT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, FilaOT].Value);
                int NumeroHoja = (int)PlanillaHojas[(int)CD.HojaNum, fila].Value;
                int ItemHoja = (int)PlanillaHojas[(int)CD.Item, fila].Value;
                double HorasEstimadas = (double)PlanillaHojas[(int)CD.HorasEstim, fila].Value;

                Actualizacion pl = new()
                {
                    Numero_OT = Numero_OT,
                    FilaOTActualizada = FilaOT,
                    NumeroHoja = NumeroHoja,
                    ItemHoja = ItemHoja,
                    HorasEstimadas = HorasEstimadas,
                    ActualizacionFechaItem = true,
                    Actualizar = Actualizacion.TipoActualizacion.AumentoHoras
                };
                pl.ShowDialog(this);
                if (pl.ActualizacionCorrecta)
                {
                    if (ListarHojasOT(FilaOT))
                    {                        
                        ColorProceso(PlanillaHojas);
                    }
                    PlanillaHojas.CurrentCell = PlanillaHojas[0, fila];
                }
            }
        }
        private void ActualizaFechasProceso(bool diasCorridos)
        {
            //Actualiza fecha de inicio-fin sólo del proceso 'En espera' seleccionado            
            int fila = PlanillaHojas.CurrentRow.Index;
            int FilaOT = PlanillaOT.CurrentCell.RowIndex;
            int Numero_OT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, FilaOT].Value);
            int NumeroHoja = (int)PlanillaHojas[(int)CD.HojaNum, fila].Value;
            int ItemHoja = (int)PlanillaHojas[(int)CD.Item, fila].Value;
            double.TryParse(PlanillaHojas[(int)CD.HorasEstim, fila].Value.ToString(), out double horasEstimadas);
            double.TryParse(PlanillaHojas[(int)CD.HorasCargadas, fila].Value.ToString(), out double horasCargadas);
            Actualizacion pl = new()
            {
                Numero_OT = Numero_OT,
                FilaOTActualizada = FilaOT,
                NumeroHoja = NumeroHoja,
                ItemHoja = ItemHoja,
                HorasEstimadas = horasEstimadas - horasCargadas,
                ActualizacionFechaItem = true,
                Actualizar = Actualizacion.TipoActualizacion.FechaInicioFinProceso,
                DiasCorridos = diasCorridos
            };
            pl.ShowDialog(this);
            if (pl.ActualizacionCorrecta)
            {
                ListarDatos();
                PlanillaOT.CurrentCell = PlanillaOT[0, FilaOT];
                PlanillaHojas.CurrentCell = PlanillaHojas[0, fila];
            }
        }
        private void ActualizaFechasIniFinProceso(bool diasCorridos, Actualizacion.TipoActualizacion tipo)
        {
            //Actualiza fecha de inicio-fin de todos los procesos 'En espera' siguientes al seleccionado           
            int fila = PlanillaHojas.CurrentRow.Index;
            int FilaOT = PlanillaOT.CurrentCell.RowIndex;
            int Numero_OT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, FilaOT].Value);
            int NumeroHoja = (int)PlanillaHojas[(int)CD.HojaNum, fila].Value;
            int ItemHoja = (int)PlanillaHojas[(int)CD.Item, fila].Value;
            double.TryParse(PlanillaHojas[(int)CD.HorasEstim, fila].Value.ToString(), out double horasEstimadas);
            double.TryParse(PlanillaHojas[(int)CD.HorasCargadas, fila].Value.ToString(), out double horasCargadas);
            Actualizacion pl = new()
            {
                Numero_OT = Numero_OT,
                FilaOTActualizada = FilaOT,
                NumeroHoja = NumeroHoja,
                ItemHoja = ItemHoja,
                ActualizacionFechaItem = true,
                DiasCorridos = diasCorridos,
                HorasEstimadas = horasEstimadas - horasCargadas,
                Actualizar = tipo
            };
            pl.ShowDialog(this);
            if (pl.ActualizacionCorrecta)
            {
                Funciones.Mensajes(PanelMsg, pl.Mensaje, esError: false);
                ListarDatos();
                PlanillaOT.CurrentCell = PlanillaOT[0, FilaOT];
                PlanillaHojas.CurrentCell = PlanillaHojas[0, fila];
            }
            else
                SetErrorMsg(pl.Mensaje);            
        }
        private void ActualizaEstadoProceso()
        {           
            int fila = PlanillaHojas.CurrentRow.Index;

            int Numero_OT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, fila].Value);
            int NumeroHoja = (int)PlanillaHojas[(int)CD.HojaNum, fila].Value;
            int ItemHoja = (int)PlanillaHojas[(int)CD.Item, fila].Value;
            var pl = new Actualizacion()
            {
                Numero_OT = Numero_OT,
                NumeroHoja = NumeroHoja,
                ItemHoja = ItemHoja,
                Actualizar = Actualizacion.TipoActualizacion.EstadoProc                
            };                        
            pl.ShowDialog(this);
            if (pl.ActualizacionCorrecta)
            {
                ListarDatos();
                PlanillaHojas.CurrentCell = PlanillaHojas[0, fila];
            }
        }       
        private void ActualizarSeguimiento(DataGridView flex)
        {
            try
            {
                string msg = string.Empty;
                int fila = PlanillaOT.CurrentRow.Index;
                int.TryParse(PlanillaOT[(int)CF.NumeroOT, fila].Value.ToString(), out int numero);
                var pl = new ActualizacionesPlanificacion()
                {
                    NumeroOT = numero,
                };
                bool result;
                flex.CommitEdit(DataGridViewDataErrorContexts.Commit);
                if (flex[1, flex.RowCount - 1].Value == null)
                {
                    SetErrorMsg("El valor de la celda no es válido. Presione enter para asegurar valor correcto");
                    return;
                }
                string coment = flex[1, flex.RowCount - 1].Value.ToString();
                var tipo = flex.Name switch
                {
                    "GrillaSegProc" => TipoSeguimiento.TiposDeSeguimiento.ProcesosOT,
                    "GrillaSegRepInter" => TipoSeguimiento.TiposDeSeguimiento.RepuestosInternacionales,
                    "GrillaSegRepNac" => TipoSeguimiento.TiposDeSeguimiento.RepuestosNacionales,
                    "GrillaSegMat" => TipoSeguimiento.TiposDeSeguimiento.Materiales,
                    "GrillaSegSub" => TipoSeguimiento.TiposDeSeguimiento.Subcontratos,
                    "GrillaSegAporte" => TipoSeguimiento.TiposDeSeguimiento.AporteCliente,
                    _ => TipoSeguimiento.TiposDeSeguimiento.ProcesosOT,
                };
                result = pl.ActualizarSeguimiento(tipo, coment, ref msg);
                if (result)
                {
                    Funciones.SetMensajeInfo(PanelMsg, "Seguimiento actualizado");
                    Funciones.ToolStripStatus = PanelMsg;
                    Funciones.DuracionMensaje();
                }
                else
                    SetErrorMsg(msg);
            }
            catch (Exception ex)
            {
                SetErrorMsg(ex.Message);
            }            
        }       
        private void ActualizarProcesosParalelos()
        {
            var plan = new ActualizacionesPlanificacion();
            int fila = PlanillaOT.CurrentRow.Index;
            int filaHoja = PlanillaHojas.CurrentRow.Index;
            plan.NumeroOT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, fila].Value);
            plan.NumeroHojaRuta = int.Parse(PlanillaHojas[(int)CD.HojaNum, filaHoja].Value.ToString());
            var lista = ObtenerProcesosParalelos(int.Parse(PlanillaHojas[(int)CD.HojaNum, filaHoja].Value.ToString()));            
            if (plan.ActualizaSecuenciaProcesosParalelos(lista))
                Funciones.Mensajes(PanelMsg, "Secuencia procesos actualizada", esError: false);
            else
                SetErrorMsg(plan.Mensaje);
            timer1.Enabled = true;
        }
        private void AccionTrabajo(byte accion)
        {
            //Detiene o reanuda orden de trabajo
            Actualizacion act = new()
            {
                //CausaSuspensionOT = PlanillaOT[(int)CF.Comentario, PlanillaOT.CurrentRow.Index].Value.ToString(),
                Actualizar = Actualizacion.TipoActualizacion.SuspensionOT
            };
            if (accion == 0)
                act.Accion = Actualizacion.AccionOT.Suspender;
            else
                act.Accion = Actualizacion.AccionOT.Reanudar;

            int FilaOT = PlanillaOT.CurrentCell.RowIndex;
            int Numero_OT = Convert.ToInt32(PlanillaOT[(int)CF.NumeroOT, FilaOT].Value);
            act.Numero_OT = Numero_OT;
            act.ShowDialog(this);
            if (act.ActualizacionCorrecta)
            {
                ListarDatos();
                PlanillaOT.CurrentCell = PlanillaOT[0, FilaOT];
            }

        }

        private void ActualizarAporteCliente(int fila)
        {
            try
            {
                PlanillaOT.CommitEdit(DataGridViewDataErrorContexts.Commit);
                int numOT = int.Parse(PlanillaOT[(int)CF.NumeroOT, fila].Value.ToString());
                bool estado = bool.Parse(PlanillaOT[(int)CF.AporteClte, fila].Value.ToString());
                var plan = new ActualizacionesPlanificacion
                {
                    NumeroOT = numOT
                };
                if (plan.ActualizarAporteCliente(estado))
                    Funciones.Mensajes(PanelMsg, "Aporte cliente actualizado", esError: false);
                else
                    SetErrorMsg(plan.Mensaje);
            }
            catch (Exception ex)
            {
                SetErrorMsg(ex.Message);
            }

        }
        #endregion
        private void PorteFormulario()
        {
            int ancho = 1200;
            int alto = 700;

            Size porte = new(ancho, alto);
            this.Size = porte;
            this.CenterToScreen();
        }
        private void AbrirOT()
        {
            var reg = new RegWindows.Registro();
            string carpeta = reg.GetDirectorioRaizApps();
            string numero = PlanillaOT[(int)CF.NumeroOT, PlanillaOT.CurrentRow.Index].Value.ToString();
            var app = new AbrirAplicacion.AbrirApp("OrdenTrabajo", new string[] { numero },carpeta);
            string msg = app.Mensaje;
            if (msg != string.Empty)
                MessageBox.Show(msg);           
        }
        private void AbrirHojaRuta()
        {
            try
            {                
                var reg = new RegWindows.Registro();
                string carpeta = reg.GetDirectorioRaizApps();
                string numero = PlanillaOT[(int)CF.NumeroOT, PlanillaOT.CurrentRow.Index].Value.ToString();
                string numeroHru = PlanillaHojas[(int)CD.HojaNum, PlanillaHojas.CurrentRow.Index].Value.ToString();
                var app = new AbrirAplicacion.AbrirApp("\\HojaRuta", new string[] { numero, numeroHru }, carpeta);
                string msg = app.Mensaje;
                if (msg != string.Empty)
                    MessageBox.Show(msg);
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }        
        private void Imprimir()
        {
            string nombreReporte = string.Empty, nombreDataset = string.Empty;
            DataSet dataset = new();
            
            var pl = new Planificacion.Planificacion
            {
                Usuario = UsuarioConectado,
                MostrarOTActiva= MostrarOtActivas.Checked
            };
            
            if (GetVista == Vista.Planificacion || GetVista == Vista.Armado || GetVista == Vista.Recuperacion)
            {
                dataset = pl.GenerarDataSetImpresionPlanificacion(TablaOTs);
                nombreReporte = "rptVistaPlanificacion.rdlc";
                nombreDataset = "DataSetPlanif";               
            }
            else if (GetVista == Vista.Taller)
            {
                dataset = pl.GenerarDataSetImpresionPlanificacion(TablaOTs);
                nombreReporte = "rptVistaTaller.rdlc";
                nombreDataset = "DataSetTaller";               
            }
            else if (GetVista == Vista.Comercial)
            {
                dataset = pl.GenerarDataSetImpresionPlanificacion(TablaOTs);
                nombreReporte = "rptVistaComercial.rdlc";
                nombreDataset = "DataSetComercial";                
            }
            string periodoEntrega = "Período entre: " + string.Format("{0:dd/MM/yyyy}", FechaEntregaProduccionIni.Value)
                           + " y " + string.Format("{0:dd/MM/yyyy}", FechaEntregaProduccionFin.Value);
            var dic = new Dictionary<string, string>
            {
                { "RangoFechas", periodoEntrega }
            };
            var idoc = new ImpresionDocumentos.FormPrint
            {
                GetDataSet = dataset,
                NombreReporte = nombreReporte,
                NombreDataSet = nombreDataset,
                Parametros = dic
            };
            idoc.ShowDialog(this);
            if (idoc.Mensaje != string.Empty)
            {
                SetErrorMsg(idoc.Mensaje);
            }           
        }
        private List<string> EstadosOTSel()
        {
            List<string> ListaEstadosOT = new();

            foreach (object itemChecked in ListaEstados.CheckedItems)
            {
                var castedItem = itemChecked as AuxiliarMaestros;
                ListaEstadosOT.Add(castedItem.Codigo.ToString());
            }
            return ListaEstadosOT;
        }
        private List<string> TiposOTSel()
        {
            List<string> ListaDeTiposOT = new();
            foreach (object itemChecked in ListaTipos.CheckedItems)
            {
                var castedItem = itemChecked as AuxiliarMaestros;
                ListaDeTiposOT.Add(castedItem.Codigo.ToString());

            }
            return ListaDeTiposOT;
        }
        private void EstableceEstadosPorDefectoOTEnTaller(bool estadoMarca)
        {
            if (ListaEstados.Items.Count == 0) return;
            //En lista de estados, selecciona los utilizados por defecto
            var pl = new Planificacion.Planificacion();            
            string msg = string.Empty;            

            //Contiene los códigos de los estados de la OT         
            ListaEstadosOTTaller = pl.EstadosOtTallerOT(ref msg);
            if (ListaEstadosOTTaller != null)
            {
                for (int i = 0; i < ListaEstados.Items.Count; i++)
                {
                    ListaEstados.SetSelected(i, true);
                    int codigoEstadoLista = int.Parse(ListaEstados.SelectedValue.ToString());
                    for (int j = 0; j < ListaEstadosOTTaller.Count; j++)
                    {
                        int codigoEstado = ListaEstadosOTTaller[j].Codigo;
                        if (codigoEstado == codigoEstadoLista)
                        {
                            EventoCodigo = true;
                            ListaEstados.SetItemChecked(i, estadoMarca);                            
                            EventoCodigo = false;
                            break;
                        }
                    }
                }
                ListaEstados.SetSelected(0, true);                
            }
            else
            {                
                SetErrorMsg(msg);
            }
        }                             
        private void MostrarDetalleOT(int fila)
        {
            if (PlanillaOT[(int)CF.DetalleOT, fila].Value == null) return;
            string DetalleOT = PlanillaOT[(int)CF.DetalleOT, fila].Value.ToString();            
            TextoDetalleOT.Text = DetalleOT;
        }
        private void MergeColumnas(PaintEventArgs e, string titulo, int columnaInicio, Color color, int anchoTitulo)
        {            
            Font tipoLetra = new("Arial", 9, FontStyle.Bold);
            Pen lapiz = new(new SolidBrush(Color.Black));
            Rectangle r1 = PlanillaOT.GetCellDisplayRectangle(columnaInicio, -1, true);

            r1.Width = anchoTitulo;
            r1.Height = (r1.Height / 2) - 2;
            e.Graphics.FillRectangle(new SolidBrush(color), r1);
            StringFormat format = new()
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            if (r1.X > 0)
            {
                int posicionLineaTitulo = (r1.Y + PlanillaOT.ColumnHeadersHeight / 2) - 3;
                int anchoLineaTitulo = SumarAnchoColumnas(0, PlanillaOT.ColumnCount - 1) + PlanillaOT.RowHeadersWidth;
                e.Graphics.DrawString(titulo, tipoLetra, new SolidBrush(this.PlanillaOT.ColumnHeadersDefaultCellStyle.ForeColor), r1, format);
                e.Graphics.DrawLine(lapiz, r1.X, posicionLineaTitulo, anchoLineaTitulo, posicionLineaTitulo);
            }
                //e.Graphics.DrawString(titulo, tipoLetra, new SolidBrush(this.PlanillaOT.ColumnHeadersDefaultCellStyle.ForeColor), r1, format);            
            //Dibuja linea            
            e.Graphics.DrawLine(lapiz, r1.X - 1, r1.Y, r1.X - 1, r1.Y + PlanillaOT.Height);
        }
        private int SumarAnchoColumnas(int colIni, int colFin)
        {
            int suma = 0;
            for (int i = colIni; i <= colFin; i++)
            {
                suma += PlanillaOT.Columns[i].Width;
            }
            return suma;
        }
        private void BordesColumnas(PaintEventArgs e, int columnaInicio)
        {
            if (PlanillaOT.RowCount == 0 || GetVista != Vista.Planificacion) return;
            Rectangle r1 = this.PlanillaOT.GetCellDisplayRectangle(columnaInicio, -1, true);
            r1.X += 1;
            r1.Y += 1;
            Pen lapiz = new(new SolidBrush(Color.Black));
            e.Graphics.DrawLine(lapiz, r1.X - 2, r1.Y, r1.X - 2, r1.Y + PlanillaOT.Height);           

        }        
        private void BordesFilaPaint(PaintEventArgs e, int columnaInicio)
        {
           
            if (PlanillaOT.RowCount == 0 || GetVista != Vista.Planificacion) return;
            
            DataGridViewRow dr = PlanillaOT.CurrentRow;
            if (dr != null)
            {
                Rectangle r2 = PlanillaOT.GetCellDisplayRectangle(columnaInicio, dr.Index, true);
                Pen lapizFila = new(Color.Black, 2);
                r2.Width = PlanillaOT.Width;
                e.Graphics.DrawRectangle(lapizFila, r2);

                if (RefrescoGrilla)
                    InvalidarCeldas(dr.Index);

                RefrescoGrilla = false;
            }
           
        }               
        private void InvalidarCeldas(int fila)
        {
            for (int j = 0; j < PlanillaOT.RowCount; j++)
            {
                if (j != fila)
                    PlanillaOT.InvalidateRow(j);
            }
        }
        private void Pegar()
        {            
            //TextoComentarioOT.Text= Clipboard.GetText();
        }
        private void Copiar()
        {
            //Clipboard.SetText(TextoComentarioOT.Text);            
        }        
        private void SetErrorMsg(string mensaje)
        {
            Funciones.SetMensajeError(PanelMsg, mensaje, 2);
            Funciones.ToolStripStatus = PanelMsg;
            Funciones.DuracionMensaje();
        }
        void CargaSucursalesCliente()
        {
            var ent = new ClienteProveedor(Cliente.SelectedValue.ToString(), 1);
            ent.ListarSucursales();
            SucursalCliente.DataSource = ent.Sucursales;
            SucursalCliente.DisplayMember = "Descripcion";
            SucursalCliente.ValueMember = "Numero";
            SucursalCliente.SelectedIndex = -1;
        }
        private void TransferirComentarioSeguimiento(DataGridView flex)
        {
            string comentario = flex[1, flex.CurrentRow.Index].Value != null ? flex[1, flex.CurrentRow.Index].Value.ToString() : string.Empty;
            var frmSeg = new FormSeg(comentario)
            {
                StartPosition = FormStartPosition.CenterParent,
                TextoComentario = comentario
            };
            frmSeg.ShowDialog(this);
            if (frmSeg.TextoComentario != string.Empty)
                flex[1, flex.CurrentRow.Index].Value = frmSeg.TextoComentario;
        }
        private void FiltrarPlanificacion()
        {            
            int filasVisibles = 0;
            for (int i = 0; i < PlanillaOT.RowCount; i++)
            {                
                PlanillaOT.Rows[i].Visible = false;             
            }
            
            for (int i = 0; i < PlanillaOT.RowCount; i++)
            {
                bool.TryParse(PlanillaOT.Rows[i].Cells[(int)CF.Suspendida].Value.ToString(), out bool susp);
                bool.TryParse(PlanillaOT.Rows[i].Cells[(int)CF.AporteClte].Value.ToString(), out bool aporte);
                int.TryParse(PlanillaOT.Rows[i].Cells[(int)CF.EstadoEntrega].Value.ToString(), out int estadoEntrega);

                if (LblRetraso.Checked && estadoEntrega == 2)
                {
                    PlanillaOT.Rows[i].Visible = true;
                    filasVisibles += 1;
                }
                if (LblFuera.Checked && estadoEntrega == 3)
                {
                    PlanillaOT.Rows[i].Visible = true;
                    filasVisibles += 1;
                }
                if (LblATiempo.Checked)
                {
                    if ((estadoEntrega == 1) || (estadoEntrega == 0 && !susp))
                    {
                        PlanillaOT.Rows[i].Visible = true;
                        filasVisibles += 1;
                    }
                }                
                if (LblSuspend.Checked)
                {                    
                    if (susp)
                    {
                        PlanillaOT.Rows[i].Visible = true;
                        filasVisibles += 1;
                    }
                    
                }
                if (LblAtraso.Checked)
                {
                    if (aporte)
                    {
                        PlanillaOT.Rows[i].Visible = true;
                        filasVisibles += 1;
                    }

                }
                //if (LblHH.Checked)
                //{
                //    bool colorHH = PlanillaOT[(int)CF.HrsPendientes, i].Style.BackColor == ColorCargaHHMayorEstima;
                //    if (colorHH)
                //    {
                //        PlanillaOT.Rows[i].Visible = true;
                //        filasVisibles += 1;
                //    }
                //}

                //if (LblAtraso.Checked)
                //{
                //    bool colorFinEval = PlanillaOT[(int)CF.FechaFinEval, i].Style.BackColor == ColorAtraso;
                //    bool colorFinPlanif = PlanillaOT[(int)CF.FechaFinPlanif, i].Style.BackColor == ColorAtraso;
                //    bool colorFinPpto = PlanillaOT[(int)CF.FechaFinRep, i].Style.BackColor == ColorAtraso;
                //    bool colorFinRep = PlanillaOT[(int)CF.FechaFinRep, i].Style.BackColor == ColorAtraso;
                //    bool colorEntrega = PlanillaOT[(int)CF.FechaEntregaClte, i].Style.BackColor == ColorAtraso;
                //    if (colorFinEval || colorFinPlanif || colorFinPpto || colorFinRep || colorEntrega)
                //    {
                //        PlanillaOT.Rows[i].Visible = true;
                //        filasVisibles += 1;
                //    }
                //}
            }
            if (filasVisibles == 0)
            {
                for (int i = 0; i < PlanillaOT.RowCount; i++)
                {
                    PlanillaOT.Rows[i].Visible = true;
                    filasVisibles += 1;
                }
            }
            tsslRegistros.Text = $"{filasVisibles} órdenes de trabajo]";
        }

        private void TrabajosEnCurso()
        {
            try
            {
                int codEmp = 0, codSuc = 0, codSec = 0;
                if (EmpleadoTarea.SelectedValue != null)
                    int.TryParse(EmpleadoTarea.SelectedValue.ToString(), out codEmp);
                if (Sucursal.SelectedValue != null)
                    int.TryParse(Sucursal.SelectedValue.ToString(), out codSuc);
                if (Seccion.SelectedValue != null)
                    int.TryParse(Seccion.SelectedValue.ToString(), out codSec);
                var pl = new Planificacion.Planificacion()
                {
                    Nivel1 = codSuc,
                    Nivel5 = codSec
                };
                var fechafin = FechaTrabajo.Value.Date.AddMinutes(UnDiaMenosUnMinuto);
                var lst = pl.ListarTrabajosEnCurso(CheckTrabajosFin.Checked, FechaTrabajo.Value.Date, fechafin, codEmp);
                GrillaTrabajos.DataSource = lst;
                ConfigurarGrillaTrabajos();
            }
            catch (Exception ex)
            {
                Funciones.SetMensajeError(PanelMsg, ex.Message, ex: ex);                
            }
           
        }
        #region Capacidad                   
        private void ObtenerCapacidadesSeccion()
        {
            try
            {                
                ListarCapacidadSecciones();
                Graficar(ChartArmado, TituloEjeX, TituloGrafico, MaximoHorasSecciones);
                ListarCapacidadCentros();
                BuscaCentroEnPlanillaArmado(string.Empty);
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void ConfigurarGrillaCapacidadCentros()
        {
            PlanillaCapCentros.ScrollBars = ScrollBars.Vertical;
            PlanillaCapCentros.AllowUserToResizeColumns = true;
            PlanillaCapCentros.ColumnHeadersHeight = 40;
            PlanillaCapCentros.Columns.Add("grupo", "Centro");
            PlanillaCapCentros.Columns.Add("horas", "Capacidad Instalada");
            PlanillaCapCentros.Columns.Add("nombreseccion", "Sec");
            PlanillaCapCentros.Columns.Add("codigo", "Cod");
            PlanillaCapCentros.Columns["grupo"].Width = 200;
            PlanillaCapCentros.Columns["horas"].Width = 70;
            PlanillaCapCentros.Columns["nombreseccion"].Visible = false;
            PlanillaCapCentros.Columns["codigo"].Visible = false;
            PlanillaCapCentros.Columns["horas"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }       
        private void ConfigurarGrillaCapacidadSecciones()
        {
            PlanillaCapacidad.ScrollBars = ScrollBars.Vertical;
            PlanillaCapacidad.AllowUserToResizeColumns = true;
            PlanillaCapacidad.Columns.Add("grupo", "Sección");
            PlanillaCapacidad.Columns.Add("horas", Capacidades.CapacidadInstalada);
            PlanillaCapacidad.Columns.Add("estim", "Estimadas");
            PlanillaCapacidad.Columns.Add("carga", "Cargadas");
            PlanillaCapacidad.Columns.Add("horasporcarga", Capacidades.CapacidadPorCargar);
            PlanillaCapacidad.Columns.Add("horascarga", Capacidades.CapacidadDisponible);

            PlanillaCapacidad.Columns["grupo"].Width = 160;
            PlanillaCapacidad.Columns["horas"].Width = 60;
            PlanillaCapacidad.Columns["horascarga"].Width = 60;
            PlanillaCapacidad.Columns["horasporcarga"].Width = 65;

            PlanillaCapacidad.Columns["estim"].Width = 60;
            PlanillaCapacidad.Columns["carga"].Width = 60;

            PlanillaCapacidad.ColumnHeadersHeight = 40;
            PlanillaCapacidad.Columns["horas"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            PlanillaCapacidad.Columns["horasporcarga"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            PlanillaCapacidad.Columns["horascarga"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            PlanillaCapacidad.Columns["estim"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            PlanillaCapacidad.Columns["carga"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            PlanillaCapacidad.Columns["carga"].DefaultCellStyle.Format = "#,##.00";
            PlanillaCapacidad.Columns["estim"].DefaultCellStyle.Format = "#,##.00";
            PlanillaCapacidad.Columns["carga"].Visible = false;
            PlanillaCapacidad.Columns["estim"].Visible = false;
        }
        private void ListarCapacidadSecciones()
        {
            int codSuc = 0, codSec = 0, codNeg = 0, codProd = 0, codNiv3 = 0;
            DateTime fechaIni, fechaFin;

            fechaIni = FechaEntregaProduccionIni.Value.Date;
            fechaFin = FechaEntregaProduccionFin.Value;
            PlanillaCapacidad.RowCount = 0;

            _ = Sucursal.SelectedValue != null && int.TryParse(Sucursal.SelectedValue.ToString(), out codSuc);
            _ = Negocio.SelectedValue != null && int.TryParse(Negocio.SelectedValue.ToString(), out codNeg);
            _ = Producto.SelectedValue != null && int.TryParse(Producto.SelectedValue.ToString(), out codProd);
            _ = Seccion.SelectedValue != null && int.TryParse(Seccion.SelectedValue.ToString(), out codSec);

            if (codProd > 0)
            {
                var mae = new Maestros.Organizacion.GenNivel3();
                codNiv3 = mae.CodigoNivel3Homologado(codProd);
            }

            var capacidad = new Capacidades()
            {
                CodigoNivel1 = codSuc,
                CodigoNivel2 = codNeg,
                CodigoNivel3 = codProd,
                CodigoNivel3Gen = codNiv3,
                CodigoNivel5 = codSec,
                FechaInicial = fechaIni,
                FechaFinal = fechaFin,
                TipoDeSeccion = TipoSec,
                EstadosOT = EstadosOTSel(),
                TiposOT = TiposOTSel(),
            };
            var lst2 = capacidad.CapacidadSecciones();
                        
            if (lst2 != null)
            {
                foreach (var item in lst2)
                {
                    PlanillaCapacidad.Rows.Add(item.Seccion, item.HorasCapacidad, item.HorasEstimadas, item.HorasCargadas,
                        item.HorasDisponibles, item.HorasPorCargar
                        );                    
                }               

                MaximoValorEjeYGrafico(lst2);                
            }
            else                
                SetErrorMsg(capacidad.Mensaje);
        }        
        private void ListarCapacidadCentros()
        {
            int codSuc = 0, codSec = 0, codNeg = 0, codProd = 0;


            PlanillaCapCentros.RowCount = 0;

            _ = Sucursal.SelectedValue != null && int.TryParse(Sucursal.SelectedValue.ToString(), out codSuc);
            _ = Negocio.SelectedValue != null && int.TryParse(Negocio.SelectedValue.ToString(), out codNeg);
            _ = Producto.SelectedValue != null && int.TryParse(Producto.SelectedValue.ToString(), out codProd);
            _ = Seccion.SelectedValue != null && int.TryParse(Seccion.SelectedValue.ToString(), out codSec);

            var ct = new Capacidades()
            {
                CodigoNivel1 = codSuc,
                CodigoNivel2 = codNeg,
                CodigoNivel3 = codProd,
                CodigoNivel5 = codSec,
                FechaInicial = FechaEntregaProduccionIni.Value.Date,
                FechaFinal = FechaEntregaProduccionFin.Value,
                TipoDeSeccion = TipoSec
            };

            var lst2 = ct.CapacidadCentros();
            decimal? sumaHoras = 0;
            if (lst2 != null)
            {
                foreach (var item in lst2)
                {
                    PlanillaCapCentros.Rows.Add(item.Centro, item.HorasCapacidad, item.Seccion, item.CodigoCentro);
                    sumaHoras += item.HorasCapacidad;
                }
            }
            else                
                SetErrorMsg(ct.Mensaje);
        }
        private void BuscaCentroEnPlanillaArmado(string nombreCentro)
        {
            if (PlanillaCapCentros.RowCount > 0)
            {
                if (nombreCentro == string.Empty)
                {
                    nombreCentro = PlanillaCapCentros[2, 0].Value.ToString();
                }
                PlanillaCapCentros.ClearSelection();
                if (PlanillaCapCentros.RowCount > 0)
                {

                    for (int i = 0; i < PlanillaCapCentros.RowCount; i++)
                    {
                        PlanillaCapCentros.Rows[i].Visible = false;
                        string centro = PlanillaCapCentros[2, i].Value.ToString();
                        if (nombreCentro == centro)
                        {
                            PlanillaCapCentros.Rows[i].Visible = true;
                        }
                    }
                }
            }
        }        
        private void MaximoValorEjeYGrafico(List<Capacidades> lista)
        {
            decimal horas = 0, tmp, horasCarga = 0, tmpCarga;
            foreach (var item in lista)
            {
                if (horas > item.HorasCapacidad)
                {
                    tmp = horas;
                }
                else
                {
                    tmp = (decimal)item.HorasCapacidad;
                    horas = tmp;
                }
                if (horasCarga > item.HorasEstimadas)
                {
                    tmpCarga = horasCarga;
                }
                else
                {
                    tmpCarga = (decimal)item.HorasEstimadas;
                    horasCarga = tmpCarga;
                }

            }
            MaximoHorasSecciones = horas > horasCarga ? horas : horasCarga;
        }
        private void Graficar(Chart grafico, string tituloEjeX, string tituloGrafico, decimal maxEje)
        {
            try
            {
                var lst = new List<Graficos.DatosGrafico>();
                foreach (DataGridViewRow item in PlanillaCapacidad.Rows)
                {
                    var dats = new Graficos.DatosGrafico()
                    {
                        EjeX = item.Cells["grupo"].Value.ToString(),
                        EjeY = double.Parse(item.Cells["horas"].Value.ToString()),
                        EjeZ = double.Parse(item.Cells["horasporcarga"].Value.ToString()),
                    };
                    lst.Add(dats);
                }
                var graf = new Graficos(grafico)
                {
                    TituloEjeX = tituloEjeX,
                    TituloEjeY = "Horas",
                    TituloEjeY2 = string.Empty,
                    TituloGrafico = tituloGrafico,
                    LeyendaSerie = "Capacidad instalada",
                    LeyendaSerie2 = "Capacidad de carga",
                    MaximoEjeY = (double)maxEje,
                    ColorSerie1 = Color.YellowGreen,
                    ColorSerie2 = Color.Red
                };

                graf.CrearGrafico(lst);
            }
            catch (Exception ex)
            {                
                SetErrorMsg(ex.Message);
            }
        }
        private void NumeroSemanaDesdeFecha()
        {
            EtiqSem1.Text = "Semana " + ISOWeek.GetWeekNumber(FechaEntregaProduccionIni.Value).ToString();
            EtiqSem2.Text = "Semana " + ISOWeek.GetWeekNumber(FechaEntregaProduccionFin.Value).ToString();

        }
        private class Ausencias
        {
            internal int Codigo { get; set; }
            internal string Letra { get; set; }
        }
        
        private void ImprimirGrafico()
        {
            PrintingManager prt = ChartArmado.Printing;
            prt.Print(true);
        }
        #endregion        

        #region Eventos
        protected async override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            var plan = new Planificacion.Planificacion();            
            var reg = new RegWindows.Registro();
            UsuarioConectado = reg.GetUsuarioConectadoRegistro();
            plan.Usuario = UsuarioConectado;
            RolUsuarioConectado = plan.RolUsuario();

            InicializarValoresControles();
            ConfigurarGrillas();
            ConfigurarGrillaCapacidadSecciones();
            ConfigurarGrillaCapacidadCentros();            
            ConfigurarGrillaPlanificacion();            
            TextoEncabezadoPlanillaOT();
            HabilitarEdicionCeldas();
            AnchoColumnasPlanillaOT();
            FormatoFechasPlanilla();
            ConfigurarGrillaHojasOT();
            ConfigurarGrillaSeguimientos(GrillaSegProc);
            ConfigurarGrillaSeguimientos(GrillaSegRepInter);
            ConfigurarGrillaSeguimientos(GrillaSegRepNac);
            ConfigurarGrillaSeguimientos(GrillaSegMat);
            ConfigurarGrillaSeguimientos(GrillaSegSub);
            ConfigurarGrillaSeguimientos(GrillaSegAporte);

            await InicializarObjetos();
            PorteFormulario();
            int codigoNivel1 = plan.EstableceSucursalUsuario();
            if (codigoNivel1 > 0)
            {
                chkNivel1.Checked = true;
                chkSeccion.Enabled = true;
                Sucursal.SelectedValue = codigoNivel1;
                CargaNegocios();
            }
            EstablecerMesActual();
            OpcionesPorVista();
            OcultarColumnasSegunVista();
            EstableceEstadosPorDefectoOTEnTaller(true);
        }
        private void BotonEjecutar_Click(object sender, EventArgs e)
        {
            if (ListarDatos())
            {
                ObtenerCapacidadesSeccion();
                if (MenuVentas.Checked && (GetVista == Vista.Planificacion || GetVista == Vista.Comercial))
                {
                    Cursor = Cursors.WaitCursor;
                    ObtenerDatosVentaSAP();
                    AgregarVendedorCombo();
                    Cursor = Cursors.Default;
                }
            }        
        }
        
        private void TsbSalir_Click(object sender, EventArgs e)
        {
            
            Close();
        }
        private void SegúnFechaEntregaClienteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Imprimir();
        }

        private void SegunFechaEntregaProducciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Imprimir();
        }
        private void ChkNivel1_CheckedChanged(object sender, EventArgs e)
        {
            if (EventoCodigo == false)
            {
                HabilitarControles(Sucursal, chkNivel1.Checked);
                chkNivel2.Checked = false;
                chkNivel3.Checked = false;
                chkSeccion.Checked = false;                
            }                
        }        

        private void ChkNivel3_CheckedChanged(object sender, EventArgs e)
        {
            if (EventoCodigo == false)
            {
                HabilitarControles(Producto, chkNivel3.Checked);
                chkSeccion.Checked = false;
            }
                
        }
        private void ChkNivel2_CheckedChanged(object sender, EventArgs e)
        {
            if (EventoCodigo == false)
            {
                HabilitarControles(Negocio, chkNivel2.Checked);
                chkNivel3.Checked = false;
                chkSeccion.Checked = false;
            }                

            if (!chkNivel2.Checked)
            {
                CodigoNivel3Homologado = 0;               
            }
        }        

        private void ChkOTR_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarControles(NumeroOTFin, chkOTR.Checked);
            HabilitarControles(NumeroOTIni, chkOTR.Checked);
        }
        private void ChkSeccion_CheckedChanged(object sender, EventArgs e)
        {            
            if (!chkSeccion.Checked)
            {
                Seccion.SelectedIndex = -1;
            }
        }
        private void ChkResponsableOT_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarControles(ResponsableOT, chkResponsableOT.Checked);
        }
        private void ChkTiposOT_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarControles(ListaTipos, chkTiposOT.Checked);                                        
        }
        private void ChkVendedor_CheckedChanged(object sender, EventArgs e)
        {
            HabilitarControles(Vendedor, chkVendedor.Checked);

        }
        private void GrillasSeguimientos_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if ((e.ColumnIndex == dgv.Columns[1].Index) && e.Value != null)
            {
                DataGridViewCell cell = dgv.Rows[e.RowIndex].Cells[1];
                cell.ToolTipText = "Doble clic para ingresar o abrir comentario";
            }
        }

        private void PlanillaOT_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == (int)CF.AporteClte)
                ActualizarAporteCliente(e.RowIndex);
            if (e.ColumnIndex == (int)CF.InfEval || e.ColumnIndex == (int)CF.InfFinal|| e.ColumnIndex == (int)CF.InfRC)
                ActualizarEstadoInformes();
        }
        private void PlanillaOT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell celda;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                celda = PlanillaOT[e.ColumnIndex, e.RowIndex];
                celda.Style.SelectionBackColor = Color.Teal;
            }
        }
        private void PlanillaOT_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == (int)CF.NumeroOT)
                AbrirOT();
        }
        private void PlanillaOT_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (EventoCodigo == false)
                ActualizarCampos(e.RowIndex);
            else
                EventoCodigo = false;
        }
        private void PlanillaOT_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell celda;

            celda = PlanillaOT[e.ColumnIndex, e.RowIndex];
            var celdaInvisible = PlanillaOT[(int)CF.EstadoEntrega, e.RowIndex];
            celda.Style.SelectionBackColor = celdaInvisible.Style.SelectionBackColor;
        }
        private void PlanillaOT_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.ColumnIndex >= 0)
            {
                if (e.RowIndex < 0)
                    return;
                DataGridViewRow clickedRow = (sender as DataGridView).Rows[e.RowIndex];
                if (!clickedRow.Selected)
                    PlanillaOT.CurrentCell = clickedRow.Cells[e.ColumnIndex];

                var mousePosition = PlanillaOT.PointToClient(System.Windows.Forms.Cursor.Position);
                cmsPlan.Show(PlanillaOT, mousePosition);
                bool suspendida = (bool)PlanillaOT[(int)CF.Suspendida, e.RowIndex].Value;
                MenuDetenerTrabajo.Enabled = !suspendida;
                MenuReanudarTrabajo.Enabled = suspendida;

                if (RolUsuarioConectado != (int)Roles.Rol.Planif && !UsuarioTieneRolAsignado())
                {
                    MenuDetenerTrabajo.Enabled = false;
                    MenuReanudarTrabajo.Enabled = false;
                }
            }
        }
        private void PlanillaOT_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //if (PlanillaOT.CurrentCell != null && e.ColumnIndex == (int)CF.Prioridad &&
            //   !PlanillaOT.Rows[e.RowIndex].IsNewRow && e.FormattedValue.ToString() != string.Empty)
            //{
            //    bool valorCorrecto = int.TryParse(e.FormattedValue.ToString(), out int valorDigitado);
            //    if (valorCorrecto)
            //    {
            //        if (valorDigitado > ValorMaximoPrioridad)
            //        {                                                
            //            SetErrorMsg("Valor no permitido");
            //            e.Cancel = true;
            //        }
            //    }
            //    else
            //    {
            //        SetErrorMsg("Valor no permitido");
            //        e.Cancel = true;
            //    }
            //}
            //else
            //    e.Cancel = false;
        }
        private void PlanillaOT_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ResaltaOTDetenida();
        }
        private void PlanillaOT_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            return;
        }
        private void PlanillaOT_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (EventoCodigo == true) return;
            if (ListarHojasOT(e.RowIndex))
            {
                ColorProceso(PlanillaHojas);
                MostrarDetalleOT(e.RowIndex);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegProc, TipoSeguimiento.TiposDeSeguimiento.ProcesosOT);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegRepInter, TipoSeguimiento.TiposDeSeguimiento.RepuestosInternacionales);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegRepNac, TipoSeguimiento.TiposDeSeguimiento.RepuestosNacionales);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegMat, TipoSeguimiento.TiposDeSeguimiento.Materiales);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegSub, TipoSeguimiento.TiposDeSeguimiento.Subcontratos);
                ObtenerSeguimientoProcesos(e.RowIndex, GrillaSegAporte, TipoSeguimiento.TiposDeSeguimiento.AporteCliente);
            }
        }
        private void PlanillaOT_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rtHeader = this.PlanillaOT.DisplayRectangle;
            rtHeader.Height = this.PlanillaOT.ColumnHeadersHeight / 2;
            this.PlanillaOT.Invalidate(rtHeader);
        }
        private void PlanillaOT_Scroll(object sender, ScrollEventArgs e)
        {
            Rectangle rtHeader = this.PlanillaOT.DisplayRectangle;
            rtHeader.Height = this.PlanillaOT.ColumnHeadersHeight / 2;
            this.PlanillaOT.Invalidate(rtHeader);

        }
        private void PlanillaOT_Paint(object sender, PaintEventArgs e)
        {
            if (GetVista == Vista.Planificacion || GetVista == Vista.Armado || GetVista == Vista.Recuperacion)
            {
                MergeColumnas(e, "Tiempo", (int)CF.HrsEstimadas, Color.FromArgb(200, 200, 200), SumarAnchoColumnas((int)CF.HrsEstimadas, (int)CF.HrsPendientes));
                MergeColumnas(e, "Dias en taller", (int)CF.FechaRecepcion, Color.FromArgb(200, 200, 200), SumarAnchoColumnas((int)CF.FechaRecepcion, (int)CF.DiasEnProceso));
                MergeColumnas(e, "Plazos de Evaluación", (int)CF.FechaInicioEval, Color.FromArgb(200, 200, 200), SumarAnchoColumnas((int)CF.FechaInicioEval, (int)CF.FechaReprogEval));
                MergeColumnas(e, "Plazos de Planificacion", (int)CF.FechaInicioPlanif, Color.FromArgb(190, 190, 190), SumarAnchoColumnas((int)CF.FechaInicioPlanif, (int)CF.FechaReprogPlanif));
                MergeColumnas(e, "Plazos de Presupuesto", (int)CF.FechaInicioPpto, Color.FromArgb(180, 180, 180), SumarAnchoColumnas((int)CF.FechaInicioPpto, (int)CF.FechaReprogPpto));
                MergeColumnas(e, "Plazos de Reparación", (int)CF.FechaInicioRep, Color.FromArgb(170, 170, 170), SumarAnchoColumnas((int)CF.FechaInicioRep, (int)CF.FechaArribo));
                MergeColumnas(e, "Plazos de Entrega", (int)CF.FechaEntregaProd, Color.FromArgb(160, 160, 160), SumarAnchoColumnas((int)CF.FechaEntregaProd, (int)CF.FechaEntregaClte));
                MergeColumnas(e, "Informes", (int)CF.InfRC, Color.FromArgb(140, 140, 140), SumarAnchoColumnas((int)CF.InfRC, (int)CF.InfFinal));
                MergeColumnas(e, "Datos SAP venta", (int)CF.Vendedor, Color.FromArgb(150, 150, 150), SumarAnchoColumnas((int)CF.Vendedor, (int)CF.OfertaVenta));

                BordesColumnas(e, (int)CF.Avance);
                BordesFilaPaint(e, 0);
            }
        }
        private void PlanillaOT_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height / 2;
                r2.Height = e.CellBounds.Height / 2;
                e.PaintBackground(r2, true);
                e.PaintContent(r2);
                e.Handled = true;

            }
        }
        

        private void PlanillaHojas_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.ColumnIndex >= 0)
            {
                
                var mousePosition = PlanillaHojas.PointToClient(System.Windows.Forms.Cursor.Position);
                cmsHojas.Show(PlanillaHojas, mousePosition);                
                MenuAbrirHojaRuta.Enabled = true;

                if (RolUsuarioConectado != (int)Roles.Rol.Planif)
                {
                    MenuActualFechasProcesos.Enabled = false;
                    MenuActualizarFechaProcTodo.Enabled = false;                                                          

                    MenuActualizarHorasEstimadasProceso.Enabled = false;
                    MenuActualizarEstadoProceso.Enabled = false;
                    MenuActualizarProcParal.Enabled = false;
                }
                else
                {                    
                    int EstadoProc = int.Parse(PlanillaHojas[(int)CD.CodigoEstadoProc, e.RowIndex].Value.ToString());
                    string Hoja = PlanillaHojas[(int)CD.HojaNum, e.RowIndex].Value.ToString();                    
                    // Actualizar fecha inicio todos los items 'En espera'
                    MenuActualizarFechaInicioTodosHabiles.Text = "Hoja " + Hoja + " - Dias hábiles";
                    MenuActualizarFechaInicioTodosCorridos.Text = "Hoja " + Hoja + " - Dias corridos";                   

                    MenuActualFechasProcesos.Enabled = EstadoProc == (int)EstadoProcesoProduccion.Estados.Espera ||
                        EstadoProc == (int)EstadoProcesoProduccion.Estados.Ejecucion;                                        
                    MenuActualizarHorasEstimadasProceso.Enabled = EstadoProc == (int)EstadoProcesoProduccion.Estados.Ejecucion;
                    MenuActualizarEstadoProceso.Enabled = (EstadoProc == (int)EstadoProcesoProduccion.Estados.Espera
                        || EstadoProc == (int)EstadoProcesoProduccion.Estados.Ejecucion);
                }
            }
        }
        private void Sucursal_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Sucursal.SelectedIndex >= 0)
            {
                EventoCodigo = true;
                chkNivel1.Checked = true;
                EventoCodigo = false;
                CargaNegocios();
            }
           
        }
        private void Negocio_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Negocio.SelectedIndex >= 0)
            {
                EventoCodigo = true;
                chkNivel2.Checked = true;
                EventoCodigo = false;
                CargaNivel3();
            }                
        }
        private void Producto_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Producto.SelectedIndex >= 0)
            {
                EventoCodigo = true;
                chkNivel3.Checked = true;
                EventoCodigo = false;
                CargaNivel5();
            }
            
        }
        private void Seccion_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Seccion.SelectedIndex >= 0)
            {
                EventoCodigo = true;
                chkSeccion.Checked = true;
                EventoCodigo = false;                
            }
        }       
        private void MenuAbrirRutaHoja_Click(object sender, EventArgs e)
        {
            AbrirHojaRuta();
        }
        private void MenuAbrirHojaRuta_Click(object sender, EventArgs e)
        {
            AbrirHojaRuta();
        }
        private void MenuActualizarFechaInicioCorridos_Click(object sender, EventArgs e)
        {
            ActualizaFechasProceso(true);
        }
        private void MenuActualizarFechaInicioHabiles_Click(object sender, EventArgs e)
        {
            ActualizaFechasProceso(false);
        }
        private void MenuActualizarFechaInicioTodosHabiles_Click(object sender, EventArgs e)
        {
            ActualizaFechasIniFinProceso(false, Actualizacion.TipoActualizacion.FechaInicioFinProcesosHoja);
        }
        private void MenuActualizarFechaInicioTodosCorridos_Click(object sender, EventArgs e)
        {
            ActualizaFechasIniFinProceso(true, Actualizacion.TipoActualizacion.FechaInicioFinProcesosHoja);
        }       
        private void MenuActualFechasProcesoHojas_Click(object sender, EventArgs e)
        {
            ActualizaFechasIniFinProceso(false, Actualizacion.TipoActualizacion.FechaInicioFinProcesosOT);
        }
        private void MenuActualFechasProcesoHojasCorr_Click(object sender, EventArgs e)
        {
            ActualizaFechasIniFinProceso(true, Actualizacion.TipoActualizacion.FechaInicioFinProcesosOT);
        }
        private void MenuActualizarHorasEstimadasProceso_Click(object sender, EventArgs e)
        {
            ActualizarHHEstimadas();
        }       
        private void MenuActualizarEstadoProceso_Click(object sender, EventArgs e)
        {
            ActualizaEstadoProceso();
        }
        private void MenuReanudarTrabajo_Click(object sender, EventArgs e)
        {
            AccionTrabajo(1);
        }

        private void MenuDetenerTrabajo_Click(object sender, EventArgs e)
        {
            AccionTrabajo(0);            
        }

        private void MenuActualizarProcParal_Click(object sender, EventArgs e)
        {            
            ActualizarProcesosParalelos();            
        }
        private void MenuOcultarColumna_Click(object sender, EventArgs e)
        {
            int col = PlanillaOT.CurrentCell.ColumnIndex;
            PlanillaOT.Columns[col].Visible = false;
        }
        private void PrintPlanificacionProduccion_Click(object sender, EventArgs e)
        {
            Imprimir();
        }        
        private void MenuAbrirOT_Click(object sender, EventArgs e)
        {
            AbrirOT();
        }
       
        private void Timer1_Tick(object sender, EventArgs e)
        {
            PanelMsg.Text = string.Empty;
            timer1.Enabled = false;
        }                                 
        private void EstadosTaller_CheckedChanged(object sender, EventArgs e)
        {
            EstableceEstadosPorDefectoOTEnTaller(EstadosTaller.Checked);                  
        }                                   
        private void FechaEntregaClienteCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FechaEntregaClienteCheck.Checked)
            {
                FechaEntregaProduccionCheck.Checked = false;                
                FechaEntregaProduccionIni.Enabled = false;
                FechaEntregaProduccionFin.Enabled = false;               
                FechaEntregaIni.Enabled = true;
                FechaEntregaFin.Enabled = true;
            }               
        }
        private void FechaEntregaProduccionCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FechaEntregaProduccionCheck.Checked)
            {

                FechaEntregaClienteCheck.Checked = false;                              
                FechaEntregaIni.Enabled = false;
                FechaEntregaFin.Enabled = false;                
                FechaEntregaProduccionIni.Enabled = true;
                FechaEntregaProduccionFin.Enabled = true;
            }                
        }        
        private void NumeroOTIni_TextChanged(object sender, EventArgs e)
        {
            NumeroOTFin.Text = NumeroOTIni.Text;
            
        }
        private void NumeroOTIni_Leave(object sender, EventArgs e)
        {
            //chkOTR.Checked = true;
        }
        private void NumeroOTIni_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != TeclaBackSpace;
        }
        private void NumeroOTFin_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != TeclaBackSpace;
        }                
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape)
            {
                EventoCodigo = true;
                PanelMsg.Text = string.Empty;               
                EventoCodigo = false;
            }
            else if (e.KeyCode == Keys.F7)
            {
                if (MostrarOcultarSeguimientos)
                {
                    panel2.BringToFront();
                    MostrarOcultarSeguimientos = false;
                }
                else
                {
                    panel2.SendToBack();
                    MostrarOcultarSeguimientos = true;
                }                    
            }
            else if (e.Control == true && e.KeyCode == Keys.F5)
                ActualizarHHEstimadas();
            else if (e.Control == true && e.KeyCode == Keys.H)
                AbrirHojaRuta();
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                RefrescoGrilla = true;
            }
        }
        private void MenuimprimirGraf_Click(object sender, EventArgs e)
        {
            ImprimirGrafico();
        }
        private void MenuPegar_Click(object sender, EventArgs e)
        {
            Pegar();
        }
        private void MenuCopiar_Click(object sender, EventArgs e)
        {
            Copiar();
        }
        
        protected virtual void Vendedor_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FiltraPorVendedor();
        }
        private void PlanillaCapacidad_RowEnter(object sender, DataGridViewCellEventArgs e)
        {            
            BuscaCentroEnPlanillaArmado(PlanillaCapacidad[(int)ColCap.Seccion, e.RowIndex].Value.ToString());
        }
        private void PlanillaCapacidadSemanal_Scroll(object sender, ScrollEventArgs e)
        {
            Rectangle rtHeader = this.PlanillaCapacidadPeriodo.DisplayRectangle;
            rtHeader.Height = this.PlanillaCapacidadPeriodo.ColumnHeadersHeight / 2;
            this.PlanillaCapacidadPeriodo.Invalidate(rtHeader);
        }
        private void PlanillaCapacidadSemanal_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rtHeader = this.PlanillaCapacidadPeriodo.DisplayRectangle;
            rtHeader.Height = this.PlanillaCapacidadPeriodo.ColumnHeadersHeight / 2;
            this.PlanillaCapacidadPeriodo.Invalidate(rtHeader);
        }        
        private void FechaCapIni_ValueChanged(object sender, EventArgs e)
        {
            NumeroSemanaDesdeFecha();
        }
        private void FechaCapFin_ValueChanged(object sender, EventArgs e)
        {
            NumeroSemanaDesdeFecha();
        }
        private void MenuFecEntrEval_Click(object sender, EventArgs e)
        {
            MenuFecEntrEval.Checked = !MenuFecEntrEval.Checked;
        }
        private void MenuMostrarHH_Click(object sender, EventArgs e)
        {
            MenuMostrarHH.Checked = !MenuMostrarHH.Checked;
            PlanillaCapacidad.Columns["carga"].Visible = MenuMostrarHH.Checked;
            PlanillaCapacidad.Columns["estim"].Visible = MenuMostrarHH.Checked;
        }
        private void MenuVentas_Click(object sender, EventArgs e)
        {
            MenuVentas.Checked = !MenuVentas.Checked;
        }
        private void GrillasSeguimientos_DoubleClick(object sender, EventArgs e)
        {
            TransferirComentarioSeguimiento((DataGridView)sender);
        }        
        private void MenuActualizarSegProc_Click(object sender, EventArgs e)
        {
            Type ty = this.ActiveControl.GetType();
            if (ty.Equals(typeof(DataGridView)))
                ActualizarSeguimiento((DataGridView)this.ActiveControl);
        }

        private void ChkSucClte_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkSucClte.Checked)
            {
                SucursalCliente.SelectedIndex = -1;
            }
        }

        private void SucursalCliente_SelectionChangeCommitted(object sender, EventArgs e)
        {
            chkSucClte.Checked = true;
        }

        private void ChkClte_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkClte.Checked)
            {
                Cliente.SelectedIndex = -1;
                SucursalCliente.SelectedIndex = -1;
            }

        }

        private void Cliente_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CargaSucursalesCliente();
            chkClte.Checked = true;
        }
        private void LblATiempo_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }

        private void LblHH_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }

        private void LblRetraso_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }

        private void LblFuera_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }

        private void LblAtraso_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }

        private void LblSuspend_CheckedChanged(object sender, EventArgs e)
        {
            FiltrarPlanificacion();
        }
        #endregion        
    }

}