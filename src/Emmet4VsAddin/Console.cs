using System.Diagnostics;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;

namespace Emmet4VsAddin
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Console
    {
        private DTE2 _dte;
        private AddIn _addIn;
        private ScriptEngine _engine;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dte"></param>
        /// <param name="addIn"></param>
        /// <param name="engine"></param>
        /// <returns></returns>
        public Console(DTE2 dte, AddIn addIn, ScriptEngine engine)
        {
            _dte = dte;
            _addIn = addIn;
            _engine = engine;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public void log(string text)
        {
            Debug.Print(text);
        }
    }
}
