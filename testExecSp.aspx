<%@ Page LANGUAGE="C#" explicit="true" Codepage="65001" %>
<%
StoreProcedure a;
string sReturn;
string StoreName;
string StrParameter;
string StrParameterOutput;
a = new StoreProcedure("cf_hbv"); //-- Key
StoreName = "SP_CMS__Company_List";
StrParameter = "@CompanyID;3;1;4;-1$@Parent;3;1;4;-1$@Status;2;1;2;-1$@Page;3;1;4;1$@PageSize;3;1;4;100$@Rowcount;3;3;4;0";
StrParameterOutput = "";
sReturn = a.ExecuteStore (StoreName, StrParameter, ref StrParameterOutput);
Response.Write ("<hr>sReturn=" + sReturn);
Response.Write ("<hr>StrParameterOutput=" + StrParameterOutput);
%>