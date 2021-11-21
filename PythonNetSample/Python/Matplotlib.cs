// ref: https://github.com/yagweb/pythonnetLab/blob/master/pynetLab/Matplotlib.cs
using System;

namespace Python.Runtime
{
    public class Matplotlib
    {
        public static void Initialize(string backend = null)
        {
            np = Py.Import("numpy");
            matplotlib = Py.Import("matplotlib");
            if (!String.IsNullOrEmpty(backend))
            {
                // backend must be set before pylab be imported
                // It can not be changed after pylab be imported
                PythonEngine.Exec($"import matplotlib;matplotlib.use('{backend}')");
            }
            plt = Py.Import("matplotlib.pylab");
            PltFigureType = plt.GetAttr("Figure").Handle;
            var io = Py.Import("io");
            BytesIO = io.GetAttr("BytesIO");
            PythonEngine.Exec("import matplotlib;print(matplotlib.get_backend())");
            PythonEngine.Exec("from matplotlib import rcParams;" +
                              "rcParams['font.family'] = 'sans-serif';" +
                              "rcParams['font.sans-serif'] = ['Hiragino Maru Gothic Pro','Yu Gothic','Meirio','Takao','IPAexGothic','IPAPGothic','VL PGothic','Noto Sans CJK JP']");
        }

        internal static PyObject np;

        public static PyObject matplotlib;

        public static PyObject plt;

        internal static IntPtr PltFigureType;

        internal static PyObject BytesIO;

        public static byte[] SaveFigureToArray(PyObject fig, int dpi = 200, string format = "png")
        {
            if (fig.GetPythonType().Handle != PltFigureType)
            {
                throw new Exception("object is not a matplotlib Figure");
            }

            dynamic _np = np;
            //buf = io.BytesIO()
            dynamic buf = BytesIO.Invoke();
            //fig.savefig(buf, dpi=__dpi__, format='png')
            fig.InvokeMethod("savefig", new PyTuple(new PyObject[] { buf }), Py.kw("dpi", dpi, "format", format));
            var buf_out = _np.array(buf.getbuffer(), Py.kw("dtype", Numpy.GetNumpyDataType(typeof(byte))));
            var arr = Numpy.ToArray(buf_out);
            return (byte[])arr;
        }

        public static string get_backend()
        {
            return matplotlib.InvokeMethod("get_backend").ToString();
        }
    }
}
