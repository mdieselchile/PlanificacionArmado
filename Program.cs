using Recursos;
using System;
using System.Windows.Forms;

namespace PlanificacionArmado
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {            
            string msg = Utilidades.BaseDatos.Parametro(Utilidades.BaseDatos.Parametros.MsgErrorAbrirApp);
            PlanArmado.Vista vista;
            if (args.Length == 0)
            {
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
            else
            {
                if (args[0] != Utilidades.BaseDatos.Parametro(Utilidades.BaseDatos.Parametros.ArgumentoAbrirApp))
                {
                    MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                    return;
                }
                else
                {
                    vista = (PlanArmado.Vista)int.Parse(args[1]);
                }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new PlanArmado(vista));
        }
    }
}
