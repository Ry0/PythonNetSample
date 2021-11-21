using Python.Runtime;
using PythonNetSample.Python;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PythonNetSample
{
    public partial class Form1 : Form
    {
        private PythonnetManager python;


        public Form1()
        {
            InitializeComponent();

            // Python呼び出し用クラスインスタンス化
            python = new PythonnetManager();
            python.PythonnetSetup();
            python.ImportLibrary();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            using (Py.GIL())
            {
                try
                {
                    var fig = python.plt.figure(Py.kw("figsize", new List<int> { 6, 3 }));
                    // -5から5まで0.1区切りで配列を作る
                    var x = python.np.arange(-5, 5, 0.1);
                    // 配列xの値に関してそれぞれsin(x)を求めてy軸の配列を生成
                    var y = python.np.sin(x);

                    python.plt.plot(x, y);
                    python.plt.show();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    python.plt.clf();
                    python.plt.close();
                }
            }
        }
    }
}
