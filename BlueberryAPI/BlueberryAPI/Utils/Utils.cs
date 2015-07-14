using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Windows.Forms;
using System.Reflection;
using Spring.Core;
using Spring.Aop.Framework;
using AopAlliance.Intercept;
using System.Net;
using System.IO;
using Dynamitey;

namespace ExcelAddIn1.Utils
{

    public interface ICommand
    {
        object Execute(object context);
    }

    public class ServiceCommand : ICommand
    {
        public object Execute(object context)
        {
            System.Diagnostics.Debug.WriteLine("Service Implementation : [{0}]", context);
            return null;
        }
    }

    public class ConsoleLoggingAroundAdvice : IMethodInterceptor
    {
        public object Invoke(IMethodInvocation invocation)
        {
            System.Diagnostics.Debug.WriteLine("Advice executing; calling the advised method...");
            object returnValue = invocation.Proceed();
            System.Diagnostics.Debug.WriteLine("Advice executed; advised method returned " + returnValue);
            return returnValue;
        }
    }

    public class CheckInternetConnection : IMethodInterceptor
    {
        public object Invoke(IMethodInvocation invocation)
        {
            /*
            try
            {
                using (var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse())
                {
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        string result = streamReader.ReadToEnd();
                        Dictionary<string, dynamic> deserializedResult = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);
                        MessageBox.Show(deserializedResult["response"]);
                    }
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse resp = ex.Response as HttpWebResponse;
                    if (resp != null && resp.StatusCode == HttpStatusCode.NotFound)
                    {
                        return;
                    }
                    else
                        throw;
                }
            }
             */
            object returnValue = invocation.Proceed();
            return returnValue;
        }
    }

    class UserManagement
    {
        public static bool userLogged()
        {
            try
            {
                var temp = GlobalVariables.sessionData["loggedUser"];
            }
            catch (Exception ex)
            {

                if (ex is NullReferenceException || ex is KeyNotFoundException)
                {
                    MessageBox.Show("Please log in");
                    return false;
                }
                throw;
                
            }
            return true;
            
        }
    }

    class BlueberryHTTPResponse
    {
        public delegate dynamic handleResponseDelegate(params object[] args);
        public delegate dynamic handleReponseExceptionsDelegate();

        private HttpWebRequest httpWebRequest;
        public HttpWebRequest HttpWebRequestParams
        {
            get { return httpWebRequest; }
            set { httpWebRequest = value; }
        }

        private dynamic data;
        public dynamic Data
        {
            get { return data; }
            set { data = value; }
        }

        private object[] handleResponseArguments;
        public object[] HandleResponseArguments
        {
            get { return handleResponseArguments; }
            set { handleResponseArguments = value; }
        }

        private Stream stream;
        public Stream StreamProperty
        {
            get { return stream; }
            set { stream = value; }
        }


        private StreamReader streamReader;
        public StreamReader StreamReaderProperty
        {
            get { return streamReader; }
            set { streamReader = value; }
        }

        private StreamWriter streamWriter;
        public StreamWriter StreamWriter
        {
            get { return streamWriter; }
            set { streamWriter = value; }
        }

        private HttpWebResponse httpResponse;
        public HttpWebResponse HttpResponse
        {
            get { return httpResponse; }
            set { httpResponse = value; }
        }

        public BlueberryHTTPResponse (HttpWebRequest httpWebRequest,
                                           dynamic data,
                                           object[] handleResponseArgument)
        {
            this.httpWebRequest = httpWebRequest;
            this.data = data;
            this.handleResponseArguments = handleResponseArgument;
        }

        public dynamic sendHTTPRequest(handleResponseDelegate handleResponse,
                                           handleReponseExceptionsDelegate handleExceptions)

        {
            try
            {
                using (stream = httpWebRequest.GetRequestStream())
                {
                    if (httpWebRequest.ContentType == "application/x-www-form-urlencoded")
                    {
                        stream.Write(data, 0, data.Length);
                        stream.Close();
                    }
                    else
                    {
                        using (streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            streamWriter.Write(data);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }
                    }
                    using (httpResponse = (HttpWebResponse)httpWebRequest.GetResponse())
                    {
                        using (streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            int handleResponseArgumentsLength = handleResponseArguments.Length;
                            object[] tempArray = new object[handleResponseArgumentsLength];
                            for (int i = 0; i < handleResponseArgumentsLength; i++)
                            {
                                var tempObject = this.GetType().GetProperty((string)handleResponseArguments[i]).GetValue(this);
                                tempArray[i] = tempObject;
                            }
                            return handleResponse(tempArray);
                        }
                    }
                }

            }
            catch (WebException ex)
            {
                if ((ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    || ex.Status == WebExceptionStatus.ConnectFailure
                    || ex.Status == WebExceptionStatus.NameResolutionFailure)
                {
                    return handleExceptions();    
                }
                else
                {
                    throw;
                }
            }
        }
    }

    class Utils
    {
        public static dynamic CallStaticMethod(string typeName, string methodName)
        {
            var type = Type.GetType(typeName);

            if (type != null)
            {
                var method = type.GetMethod(methodName);

                if (method != null)
                {
                    return method.Invoke(null, null);
                }
            }
            return "";
        }

        public static object Eval(string sCSCode)
        {

            CSharpCodeProvider c = new CSharpCodeProvider();
            ICodeCompiler icc = c.CreateCompiler();
            CompilerParameters cp = new CompilerParameters();

            cp.ReferencedAssemblies.Add("system.dll");
            cp.ReferencedAssemblies.Add("system.xml.dll");
            cp.ReferencedAssemblies.Add("system.data.dll");
            cp.ReferencedAssemblies.Add("system.windows.forms.dll");
            cp.ReferencedAssemblies.Add("system.drawing.dll");

            cp.CompilerOptions = "/t:library";
            cp.GenerateInMemory = true;

            StringBuilder sb = new StringBuilder("");
            sb.Append("using System;\n");
            sb.Append("using System.Xml;\n");
            sb.Append("using System.Data;\n");
            sb.Append("using System.Data.SqlClient;\n");
            sb.Append("using System.Windows.Forms;\n");
            sb.Append("using System.Drawing;\n");

            sb.Append("namespace CSCodeEvaler{ \n");
            sb.Append("public class CSCodeEvaler{ \n");
            sb.Append("public object EvalCode(){\n");
            sb.Append("return " + sCSCode + "; \n");
            sb.Append("} \n");
            sb.Append("} \n");
            sb.Append("}\n");

            CompilerResults cr = icc.CompileAssemblyFromSource(cp, sb.ToString());
            if (cr.Errors.Count > 0)
            {
                MessageBox.Show("ERROR: " + cr.Errors[0].ErrorText,
                   "Error evaluating cs code", MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
                return null;
            }

            System.Reflection.Assembly a = cr.CompiledAssembly;
            object o = a.CreateInstance("CSCodeEvaler.CSCodeEvaler");

            Type t = o.GetType();
            MethodInfo mi = t.GetMethod("EvalCode");

            object s = mi.Invoke(o, null);
            return s;

        }
    }
}
