﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlanificacionArmado
{
    public partial class FormSeg : Form
    {
        public FormSeg()
        {
            InitializeComponent();
            BotonOK.Click += BotonOK_Click;
            BotonCancel.Click += BotonCancel_Click;
        }
        public string TextoComentario { get; set; }

        private void BotonCancel_Click(object sender, EventArgs e)
        {
            TextoComentario = string.Empty;
            Close();
        }

        private void BotonOK_Click(object sender, EventArgs e)
        {
            TextoComentario = Comentario.Text;
            Close();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape)
            {
                TextoComentario = string.Empty;
                Close();
            }
        }

    }
}
