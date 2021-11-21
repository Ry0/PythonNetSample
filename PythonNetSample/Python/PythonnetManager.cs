using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Python.Runtime;

namespace PythonNetSample.Python
{
    public class PythonnetManager
    {
        public PythonnetManager()
        {

        }

        ~PythonnetManager()
        {
            PythonnetShutdown();
        }

        /// <summary>
        /// Python環境を登録する。
        /// </summary>
        public void PythonnetSetup()
        {
            // python環境にパスを通す
            var PYTHON_HOME = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.PYTHON_PATH);
            if (!Directory.Exists(PYTHON_HOME))
            {
                throw new DirectoryNotFoundException(PYTHON_HOME + " is not exist.");
            }

            // PATHに環境変数を追加
            AddEnvPath(PYTHON_HOME);

            // python環境に、PYTHON_HOME(標準pythonライブラリの場所)を設定
            PythonEngine.PythonHome = PYTHON_HOME;

            // python環境に、PYTHON_PATH(モジュールファイルのデフォルトの検索パス)を設定
            var pythonSitePackages = Path.Combine(PYTHON_HOME, @"Lib\site-packages");
            if (!Directory.Exists(pythonSitePackages))
            {
                throw new DirectoryNotFoundException(pythonSitePackages + " is not exist.");
            }

            PythonEngine.PythonPath = string.Join(
              Path.PathSeparator.ToString(),
              new string[] {
                  PythonEngine.PythonPath,// 元の設定を残す
                  pythonSitePackages, //pipで入れたパッケージはここに入る
              }
            );

            // 初期化
            PythonEngine.Initialize();

            // 別スレッドでの呼び出しを許可する。
            mThreadState = PythonEngine.BeginAllowThreads();

            // tcl,tkの設定
            var tclBasePath = Path.Combine(PYTHON_HOME, "tcl");
            SetTCLandTKSetting(tclBasePath);
        }

        /// <summary>
        /// Pythonでの辞書型に変換するための便利クラス
        /// </summary>
        public PyConverter Converter { set; get; }

        /// <summary>
        /// Pythonの呼び出しスレッド管理用変数
        /// </summary>
        private IntPtr mThreadState;

        public void PythonnetShutdown()
        {
            PythonEngine.EndAllowThreads(mThreadState);
            PythonEngine.Shutdown();
        }

        private dynamic np_;
        public dynamic np { get { return np_; } }

        private dynamic plt_;
        public dynamic plt { get { return plt_; } }

        /// <summary>
        /// 必要なモジュールをインポートする
        /// </summary>
        public void ImportLibrary()
        {
            using (Py.GIL())
            {
                // モジュールインポート
                Numpy.Initialize();
                np_ = Numpy.np;

                Matplotlib.Initialize("TkAgg");
                plt_ = Matplotlib.plt;

                // Pythonの辞書型に変換する用のクラス初期化
                Converter = new PyConverter();
                Converter.AddListType();
                Converter.Add(new StringType());
                Converter.Add(new Int64Type());
                Converter.Add(new Int32Type());
                Converter.Add(new FloatType());
                Converter.Add(new DoubleType());
                Converter.AddDictType<string, object>();
            }
        }

        /// <summary>
        /// プロセスの環境変数PATHに、パスを通す。
        /// </summary>
        /// <param name="paths">PATHに追加するディレクトリ。</param>
        private void AddEnvPath(params string[] paths)
        {
            SetEnvironmentVariable("PATH", paths);
        }

        /// <summary>
        /// プロセスの環境変数PATHに、指定されたディレクトリを追加する(パスを通す)。
        /// </summary>
        /// <param name="paths">PATHに追加するディレクトリ。</param>
        private void SetTCLandTKSetting(string basePath)
        {
            if (!Directory.Exists(basePath))
            {
                throw new DirectoryNotFoundException(basePath + " is not exist.");
            }

            // tcl8.6
            var tcl = Path.Combine(basePath, "tcl8.6");
            if (!Directory.Exists(tcl))
            {
                throw new DirectoryNotFoundException(tcl + " is not exist.");
            }

            // tk8.6
            var tk = Path.Combine(basePath, "tk8.6");
            if (!Directory.Exists(tk))
            {
                throw new DirectoryNotFoundException(tk + " is not exist.");
            }

            using (Py.GIL())
            {
                PythonEngine.Exec($"import os" + "\n" +
                  $"os.environ['TCL_LIBRARY'] = r'{tcl}'" + "\n" +
                  $"os.environ['TK_LIBRARY'] = r'{tk}'");
            }
        }

        /// <summary>
        /// 環境変数をセットする
        /// </summary>
        private void SetEnvironmentVariable(string variable, params string[] paths)
        {
            var envNewPaths = new List<string>();
            var envPaths = Environment.GetEnvironmentVariable(variable);
            if (envPaths != null)
            {
                envNewPaths = envPaths.Split(Path.PathSeparator).ToList();
            }

            foreach (var path in paths)
            {
                if (path.Length > 0 && !envNewPaths.Contains(path))
                {
                    envNewPaths.Insert(0, path);
                }
            }
            Environment.SetEnvironmentVariable(variable, string.Join(Path.PathSeparator.ToString(), envNewPaths), EnvironmentVariableTarget.Process);
        }
    }
}