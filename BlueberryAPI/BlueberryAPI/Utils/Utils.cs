using System;
using System.Collections.Generic;
using System.Collections.Specialized;
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
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
        public delegate dynamic handleReponseExceptionsDelegate(params object[] args);

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
                object[] exceptionArgs = new object[2] {ex.Status.ToString(), ex.Message};
                return handleExceptions(exceptionArgs);
            }
        }
    }

    class Utils
    {

        public static OrderedDictionary toOrderedDictionary (Dictionary<string, dynamic> dictionaryToBeConverted) {

            OrderedDictionary orderedDictionary = new OrderedDictionary();
            foreach (KeyValuePair<string, dynamic> kvp in dictionaryToBeConverted)
            {
                orderedDictionary.Add(kvp.Key, kvp.Value);
            }
            return orderedDictionary;
        }

        public static Dictionary<string, dynamic> toDictionary (OrderedDictionary orderedDictionaryToBeConverted) {
            Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
            List<string> keys = orderedDictionaryToBeConverted.Keys.Cast<string>().ToList();
            List<dynamic> values = orderedDictionaryToBeConverted.Values.Cast<dynamic>().ToList();
            for (int i = 0; i < orderedDictionaryToBeConverted.Keys.Count; i++)
            {
                dictionary.Add(keys[i], values[i]);
            }
            return dictionary;
        }

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

    static class ExcelCellErrors
    {
        private static int[] errors = {-2146826281,
                               -2146826246,
                               -2146826259,
                               -2146826288,
                               -2146826252,
                               -2146826265,
                               -2146826273};

        public static int[] Errors
        {
            get {return errors;}
        }

    }

    class CustomIntConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return (objectType == typeof(int));
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override void WriteJson(JsonWriter writer, dynamic value, JsonSerializer serializer)
        {

            JValue jsonValue = (JValue)value;

            if (jsonValue.Type == JTokenType.Float)
            {
                jsonValue.Value<double>();
            }
            else if (jsonValue.Type == JTokenType.Integer)
            {
                jsonValue.Value<int>();
            }

            jsonValue = serializer.Serialize(writer, value);

            throw new FormatException();
        }
    }
}
