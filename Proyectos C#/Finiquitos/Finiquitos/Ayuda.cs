using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finiquitos
{
    public partial class Ayuda : Form
    {
        public Ayuda()
        {
            InitializeComponent();
        }
        private static Ayuda m_FormDefInstance;
        public static Ayuda DefInstance
        {
            get
            {
                if (m_FormDefInstance == null || m_FormDefInstance.IsDisposed)
                    m_FormDefInstance = new Ayuda();
                return m_FormDefInstance;
            }
            set
            {
                m_FormDefInstance = value;
            }
        }
        private void Ayuda_Load(object sender, EventArgs e)
        {

        }
    }
}
