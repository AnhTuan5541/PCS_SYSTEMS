using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Reflection;
using System.Security.Claims;
using Microsoft.Data.SqlClient;
using PCS_SYSTEMS.Common;
using PCS_SYSTEMS.Response;
using System.Data.Common;
using Microsoft.AspNetCore.Authorization;
using OfficeOpenXml.Style;
using System.ComponentModel.Design;

namespace PCS_system.Controllers
{
    public class LoadingPlanController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetRawData(string userid)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);

                // GetRawData
                using var command = new SqlCommand("DA_GetRawData", connection) { CommandType = CommandType.StoredProcedure };
                command.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader = command.ExecuteReader();
                List<Dictionary<string, object>> rawData = CommonFunction.GetDataFromProcedure(reader);
                connection.Close();

                var groupedData = rawData.GroupBy(d => d["wafer_lot"])
                    .Select(g => g.ToList())
                    .ToList();

                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Get raw data successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<List<Dictionary<String, Object>>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Get raw data successfully",
                    Data = groupedData,
                    size = groupedData.Count
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        [HttpGet]
        public IActionResult GetThickness(string userid)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);
                Dictionary<string, object> dataThickness = new Dictionary<string, object>();

                // Get total die all
                using var command3 = new SqlCommand("DA_GetTotalDie", connection) { CommandType = CommandType.StoredProcedure };
                command3.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader3 = command3.ExecuteReader();
                List<Dictionary<string, object>> totalDieObject = CommonFunction.GetDataFromProcedure(reader3);
                connection.Close();
                int totalDie = Convert.ToInt32(totalDieObject[0]["total_die"]);
                dataThickness.Add("Total Die", totalDie);

                // Get total thickness and die usage
                using var command4 = new SqlCommand("DA_GetThickness", connection) { CommandType = CommandType.StoredProcedure };
                command4.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader4 = command4.ExecuteReader();
                List<Dictionary<string, object>> thickness = CommonFunction.GetDataFromProcedure(reader4);
                connection.Close();
                int totalDieUsagePerUnit = 0; // tong so luong die can dung cho 1 unit
                for (int i = 0; i < thickness.Count; i++)
                {
                    totalDieUsagePerUnit += Convert.ToInt32(thickness[i]["die_usage"]);
                }

                // Tinh so luong die can dung theo tung thichness
                for (int i = 0; i < thickness.Count; i++)
                {
                    double eachThickness = ((double)totalDie / totalDieUsagePerUnit) * Convert.ToInt32(thickness[i]["die_usage"]);
                    double roundedEachThickness = Math.Round(eachThickness, 2);
                    dataThickness.Add(thickness[i]["thickness"].ToString(), roundedEachThickness);
                }

                List<Dictionary<string, object>> listDataThickness = new List<Dictionary<string, object>>();
                listDataThickness.Add(dataThickness);

                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Get data thickness successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Login successfully",
                    Data = listDataThickness,
                    size = listDataThickness.Count()
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        [HttpGet]
        public IActionResult GetTotalDieByThickness(string userid)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);

                using var command3 = new SqlCommand("DA_GetTotalDieByThickness", connection) { CommandType = CommandType.StoredProcedure };
                command3.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader3 = command3.ExecuteReader();
                List<Dictionary<string, object>> data = CommonFunction.GetDataFromProcedure(reader3);
                connection.Close();

                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Get data thickness successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Login successfully",
                    Data = data,
                    size = data.Count()
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        [HttpPost]
        public IActionResult SaveSliceForThickness([FromBody] Dictionary<string, object> data)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            string userid = data["userid"].ToString();
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);

                using var command = new SqlCommand("DA_SaveSliceForThickness", connection) { CommandType = CommandType.StoredProcedure };
                command.Parameters.AddWithValue("@idCard", userid);
                command.Parameters.AddWithValue("@thicknessName", data["thicknessName"].ToString());
                command.Parameters.AddWithValue("@sliceIds", data["sliceIds"].ToString());
                command.Parameters.AddWithValue("@backgroundColor", data["backgroundColor"].ToString());
                connection.Open();
                var reader = command.ExecuteReader();
                List<Dictionary<string, object>> rawData = CommonFunction.GetDataFromProcedure(reader);
                connection.Close();

                using var command3 = new SqlCommand("DA_GetDieSliceByThickness", connection) { CommandType = CommandType.StoredProcedure };
                command3.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader3 = command3.ExecuteReader();
                List<Dictionary<string, object>> sliceThickness = CommonFunction.GetDataFromProcedure(reader3);
                connection.Close();

                // Get lot size
                using var command6 = new SqlCommand("DA_GetLotSize", connection) { CommandType = CommandType.StoredProcedure };
                command6.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader6 = command6.ExecuteReader();
                List<Dictionary<string, object>> lotSizeObj = CommonFunction.GetDataFromProcedure(reader6);
                int lotSize = Convert.ToInt32(lotSizeObj[0]["lot_size"]);
                connection.Close();

                //Delete slice last lot
                using var command11 = new SqlCommand("DA_DeleteSliceLastLot", connection) { CommandType = CommandType.StoredProcedure };
                command11.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader11 = command11.ExecuteReader();
                connection.Close();

                int countLastLot = 1;
                for (int i = 0; i < sliceThickness.Count; i++)
                {
                    using var command4 = new SqlCommand("DA_GetSpecByThickness", connection) { CommandType = CommandType.StoredProcedure };
                    command4.Parameters.AddWithValue("@idCard", userid);
                    command4.Parameters.AddWithValue("@thickness", sliceThickness[i]["thickness"]);
                    connection.Open();
                    var reader4 = command4.ExecuteReader();
                    List<Dictionary<string, object>> specByThickness = CommonFunction.GetDataFromProcedure(reader4);
                    connection.Close();

                    using var command2 = new SqlCommand("DA_GetSliceWithThickness", connection) { CommandType = CommandType.StoredProcedure };
                    command2.Parameters.AddWithValue("@idCard", userid);
                    command2.Parameters.AddWithValue("@thickness", sliceThickness[i]["thickness"]);
                    connection.Open();
                    var reader2 = command2.ExecuteReader();
                    List<Dictionary<string, object>> sliceList = CommonFunction.GetDataFromProcedure(reader2);
                    connection.Close();

                    int indexSliceList = 0;
                    string sliceIdClone = "";
                    string waferLotName = "";
                    int dcc = 1;

                    for (int j = 0; j< specByThickness.Count; j++)
                    {
                        int countDieEachAlloc = 0;
                        int countDieRemain = 0;
                        string dieAllocClone = specByThickness[j]["die_alloc"].ToString();
                        int dieUsage = Convert.ToInt32(specByThickness[j]["die_usage"]);
                        double lotFullDouble = (double)Convert.ToInt32(sliceThickness[i]["total_die"]) / Convert.ToInt32(sliceThickness[i]["thickness_count"]) / (lotSize * dieUsage);

                        int lotFull = Convert.ToInt32(sliceThickness[i]["total_die"]) / Convert.ToInt32(sliceThickness[i]["thickness_count"]) / (lotSize * dieUsage);
                        if (lotFullDouble - lotFull < 0.5)
                        {
                            lotFull -= 1;
                            countLastLot = 2;
                        }

                        int dieEachThickness = lotFull * lotSize * dieUsage;
                        
                        for (int k = indexSliceList; k<sliceList.Count; k++)
                        {
                            countDieEachAlloc += Convert.ToInt32(sliceList[k]["die_quantity"] == "" ? "0" : sliceList[k]["die_quantity"]);
                            countDieRemain += Convert.ToInt32(sliceList[k]["die_quantity"] == "" ? "0" : sliceList[k]["die_quantity"]);
                            sliceIdClone += sliceList[k]["id"] +";";

                            using var command20 = new SqlCommand("DA_AddDccForSlice", connection) { CommandType = CommandType.StoredProcedure };
                            command20.Parameters.AddWithValue("@id", Convert.ToInt32(sliceList[k]["id"]));
                            command20.Parameters.AddWithValue("@waferLotDcc", "");
                            connection.Open();
                            var reader20 = command20.ExecuteReader();
                            connection.Close();

                            //Tinh toan remain 
                            if (countDieRemain > lotSize * dieUsage)
                            {
                                countDieRemain = countDieRemain - lotSize * dieUsage;
                                dcc++;
                                waferLotName = sliceList[k]["wafer_lot"].ToString();
                                using var command5 = new SqlCommand("DA_AddRemainForSlice", connection) { CommandType = CommandType.StoredProcedure };
                                command5.Parameters.AddWithValue("@id", Convert.ToInt32(sliceList[k]["id"]));
                                command5.Parameters.AddWithValue("@remain", countDieRemain);
                                connection.Open();
                                var reader5 = command5.ExecuteReader();
                                connection.Close();

                            }

                            //Tinh toan cho tung die_alloc
                            if (countDieEachAlloc > dieEachThickness || k == sliceList.Count - 1)
                            {
                                using var command5 = new SqlCommand("DA_AddAllocForSlice", connection) { CommandType = CommandType.StoredProcedure };
                                command5.Parameters.AddWithValue("@sliceIds", sliceIdClone);
                                command5.Parameters.AddWithValue("@dieAlloc", dieAllocClone);
                                connection.Open();
                                var reader5 = command5.ExecuteReader();
                                connection.Close();

                                sliceIdClone = "";
                                indexSliceList = k +1;

                                using var command10 = new SqlCommand("DA_AddLastSlice", connection) { CommandType = CommandType.StoredProcedure };
                                command10.Parameters.AddWithValue("@slice_id", sliceList[k]["slice_id"]);
                                command10.Parameters.AddWithValue("@slice_number", sliceList[k]["slice_number"]);
                                command10.Parameters.AddWithValue("@die_quantity", sliceList[k]["die_quantity"]);
                                command10.Parameters.AddWithValue("@wafer_lot", sliceList[k]["wafer_lot"]);
                                command10.Parameters.AddWithValue("@id_card", sliceList[k]["id_card"]);
                                command10.Parameters.AddWithValue("@thickness", sliceList[k]["thickness"]);
                                command10.Parameters.AddWithValue("@backgroundColor", sliceList[k]["backgroundColor"]);
                                command10.Parameters.AddWithValue("@die_alloc", dieAllocClone);
                                command10.Parameters.AddWithValue("@wafer_lot_dcc", "");
                                command10.Parameters.AddWithValue("@remain", countDieRemain);
                                command10.Parameters.AddWithValue("@isFullLotRemain", 1);
                                command10.Parameters.AddWithValue("@dieEachAlloc", "");
                                command10.Parameters.AddWithValue("@toOpr", "");
                                connection.Open();
                                var reader10 = command10.ExecuteReader();
                                connection.Close();

                                break;
                            }
                        }
                    }
                }

                // Add nhung slice chua duoc assign cho alloc
                using var command9 = new SqlCommand("DA_AddSliceHasNotAlloc", connection) { CommandType = CommandType.StoredProcedure };
                command9.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader9 = command9.ExecuteReader();
                connection.Close();

                using var command22 = new SqlCommand("DA_GetThicknessWithType", connection) { CommandType = CommandType.StoredProcedure };
                command22.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader22 = command22.ExecuteReader();
                List<Dictionary<string, object>> thicknessTypeM = CommonFunction.GetDataFromProcedure(reader22);
                string typeThickness = thicknessTypeM[0]["thickness"].ToString();
                connection.Close();

                using var command12 = new SqlCommand("DA_GetThickness", connection) { CommandType = CommandType.StoredProcedure };
                command12.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader12 = command12.ExecuteReader();
                List<Dictionary<string, object>> thicknessList = CommonFunction.GetDataFromProcedure(reader12);
                connection.Close();

                if (countLastLot == 1)
                {
                    int lotSizeM = 0;
                    for (int i = 0; i < thicknessList.Count; i++)
                    {
                        using var command14 = new SqlCommand("DA_GetLastLotHasRemain", connection) { CommandType = CommandType.StoredProcedure };
                        command14.Parameters.AddWithValue("@idCard", userid);
                        command14.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader14 = command14.ExecuteReader();
                        List<Dictionary<string, object>> lastLotHasRemain = CommonFunction.GetDataFromProcedure(reader14);
                        connection.Close();

                        using var command15 = new SqlCommand("DA_GetLastLotNotHasRemain", connection) { CommandType = CommandType.StoredProcedure };
                        command15.Parameters.AddWithValue("@idCard", userid);
                        command15.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader15 = command15.ExecuteReader();
                        List<Dictionary<string, object>> lastLotNotHasRemain = CommonFunction.GetDataFromProcedure(reader15);
                        connection.Close();

                        int totalDie = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            totalDie += Convert.ToInt32(lastLotHasRemain[j]["remain"]);
                        }
                        for (int j = 0; j < lastLotNotHasRemain.Count; j++)
                        {
                            totalDie += lastLotNotHasRemain[j]["die_quantity"].ToString() == "" ? 0 : Convert.ToInt32(lastLotNotHasRemain[j]["die_quantity"]);
                        }
                        //End tinh total die

                        //if (thicknessList[i]["thickness"].ToString() == typeThickness) {
                        //    lotSizeM = totalDie;
                        //}

                        //int dieEachAlloc = lotSizeM / Convert.ToInt32(thicknessList[i]["total_layer"]) / countLastLot * Convert.ToInt32(thicknessList[i]["die_usage"]);

                        if (thicknessList[i]["thickness"].ToString() == typeThickness)
                        {
                            lotSizeM = totalDie / countLastLot / Convert.ToInt32(thicknessList[i]["total_layer"]) / Convert.ToInt32(thicknessList[i]["die_usage"]);
                        }

                        int dieEachAlloc = lotSizeM  * Convert.ToInt32(thicknessList[i]["die_usage"]);

                        int remain = 0;
                        int indexNotHasRemain = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            int totalEachAlloc = Convert.ToInt32(lastLotHasRemain[j]["remain"]) + remain;
                            int isFirst = -1;
                            for (int k = indexNotHasRemain; k < lastLotNotHasRemain.Count; k++)
                            {
                                totalEachAlloc += lastLotNotHasRemain[k]["die_quantity"].ToString() == ""? 0: Convert.ToInt32(lastLotNotHasRemain[k]["die_quantity"]);

                                using var command16 = new SqlCommand("DA_AddDieAllocLastLot", connection) { CommandType = CommandType.StoredProcedure };
                                command16.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                command16.Parameters.AddWithValue("@die_alloc", lastLotHasRemain[j]["die_alloc"]);
                                command16.Parameters.AddWithValue("@dieEachAlloc", dieEachAlloc);
                                command16.Parameters.AddWithValue("@lot_size", lotSizeM);
                                connection.Open();
                                var reader16 = command16.ExecuteReader();
                                connection.Close();

                                if (totalEachAlloc > dieEachAlloc || k == lastLotNotHasRemain.Count - 1)
                                {
                                    isFirst++;
                                    remain = totalEachAlloc - dieEachAlloc;

                                    using var command17 = new SqlCommand("DA_AddRemainLastLot", connection) { CommandType = CommandType.StoredProcedure };
                                    command17.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                    command17.Parameters.AddWithValue("@remain", remain);
                                    command17.Parameters.AddWithValue("@toOpr", j == lastLotHasRemain.Count - 1 ? "" : lastLotHasRemain[j + 1]["da_opr"]);
                                    command17.Parameters.AddWithValue("@waferLotDcc", "");

                                    connection.Open();
                                    var reader17 = command17.ExecuteReader();
                                    connection.Close();

                                    indexNotHasRemain = k + 1;
                                    break;
                                }
                            }
                        }
                    }
                }
                else if(countLastLot == 2)
                {
                    //Tinh last lot 1
                    int lotSizeM = 0;
                    for (int i = 0; i < thicknessList.Count; i++)
                    {
                        using var command14 = new SqlCommand("DA_GetLastLotHasRemain", connection) { CommandType = CommandType.StoredProcedure };
                        command14.Parameters.AddWithValue("@idCard", userid);
                        command14.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader14 = command14.ExecuteReader();
                        List<Dictionary<string, object>> lastLotHasRemain = CommonFunction.GetDataFromProcedure(reader14);
                        connection.Close();

                        using var command15 = new SqlCommand("DA_GetLastLotNotHasRemain", connection) { CommandType = CommandType.StoredProcedure };
                        command15.Parameters.AddWithValue("@idCard", userid);
                        command15.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader15 = command15.ExecuteReader();
                        List<Dictionary<string, object>> lastLotNotHasRemain = CommonFunction.GetDataFromProcedure(reader15);
                        connection.Close();

                        int totalDie = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            totalDie += Convert.ToInt32(lastLotHasRemain[j]["remain"]);
                        }
                        for (int j = 0; j < lastLotNotHasRemain.Count; j++)
                        {
                            totalDie += lastLotNotHasRemain[j]["die_quantity"].ToString() == "" ? 0 : Convert.ToInt32(lastLotNotHasRemain[j]["die_quantity"]);
                        }
                        //End tinh total die

                        //if (thicknessList[i]["thickness"].ToString() == typeThickness)
                        //{
                        //    lotSizeM = totalDie;
                        //}

                        //int dieEachAlloc = lotSizeM / Convert.ToInt32(thicknessList[i]["total_layer"]) / countLastLot * Convert.ToInt32(thicknessList[i]["die_usage"]);

                        if (thicknessList[i]["thickness"].ToString() == typeThickness)
                        {
                            lotSizeM = totalDie / countLastLot / Convert.ToInt32(thicknessList[i]["total_layer"]) / Convert.ToInt32(thicknessList[i]["die_usage"]);
                        }

                        int dieEachAlloc = lotSizeM  * Convert.ToInt32(thicknessList[i]["die_usage"]);

                        int remain = 0;
                        int indexNotHasRemain = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            int totalEachAlloc = Convert.ToInt32(lastLotHasRemain[j]["remain"]);
                            int isFirst = -1;
                            for (int k = indexNotHasRemain; k < lastLotNotHasRemain.Count; k++)
                            {
                                totalEachAlloc += lastLotNotHasRemain[k]["die_quantity"].ToString() == "" ? 0 : Convert.ToInt32(lastLotNotHasRemain[k]["die_quantity"]);

                                using var command16 = new SqlCommand("DA_AddDieAllocLastLot", connection) { CommandType = CommandType.StoredProcedure };
                                command16.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                command16.Parameters.AddWithValue("@die_alloc", lastLotHasRemain[j]["die_alloc"]);
                                command16.Parameters.AddWithValue("@dieEachAlloc", dieEachAlloc);
                                command16.Parameters.AddWithValue("@lot_size", lotSizeM);
                                connection.Open();
                                var reader16 = command16.ExecuteReader();
                                connection.Close();

                                if (totalEachAlloc > dieEachAlloc || k == lastLotNotHasRemain.Count - 1)
                                {
                                    isFirst++;
                                    remain = totalEachAlloc - dieEachAlloc;

                                    using var command17 = new SqlCommand("DA_AddRemainLastLot", connection) { CommandType = CommandType.StoredProcedure };
                                    command17.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                    command17.Parameters.AddWithValue("@remain", remain);
                                    command17.Parameters.AddWithValue("@toOpr", lastLotHasRemain[j]["da_opr"]);
                                    command17.Parameters.AddWithValue("@waferLotDcc", "");

                                    connection.Open();
                                    var reader17 = command17.ExecuteReader();
                                    connection.Close();

                                    using var command10 = new SqlCommand("DA_AddLastSlice2", connection) { CommandType = CommandType.StoredProcedure };
                                    command10.Parameters.AddWithValue("@slice_id", lastLotNotHasRemain[k]["slice_id"]);
                                    command10.Parameters.AddWithValue("@slice_number", lastLotNotHasRemain[k]["slice_number"]);
                                    command10.Parameters.AddWithValue("@die_quantity", lastLotNotHasRemain[k]["die_quantity"]);
                                    command10.Parameters.AddWithValue("@wafer_lot", lastLotNotHasRemain[k]["wafer_lot"]);
                                    command10.Parameters.AddWithValue("@id_card", lastLotNotHasRemain[k]["id_card"]);
                                    command10.Parameters.AddWithValue("@thickness", lastLotNotHasRemain[k]["thickness"]);
                                    command10.Parameters.AddWithValue("@backgroundColor", lastLotNotHasRemain[k]["backgroundColor"]);
                                    command10.Parameters.AddWithValue("@die_alloc", lastLotHasRemain[j]["die_alloc"]);
                                    command10.Parameters.AddWithValue("@wafer_lot_dcc", lastLotNotHasRemain[k]["wafer_lot_dcc"]);
                                    command10.Parameters.AddWithValue("@remain", remain);
                                    command10.Parameters.AddWithValue("@isFullLotRemain", 1);
                                    command10.Parameters.AddWithValue("@dieEachAlloc", "");
                                    command10.Parameters.AddWithValue("@toOpr", "");
                                    connection.Open();
                                    var reader10 = command10.ExecuteReader();
                                    connection.Close();

                                    indexNotHasRemain = k + 1;
                                    break;
                                }
                            }
                        }
                    }

                    //Tinh cho last lot 2
                    // Add nhung slice vào last lot 2 chua duoc assign cho alloc
                    using var command18 = new SqlCommand("DA_AddSliceHasNotAlloc2", connection) { CommandType = CommandType.StoredProcedure };
                    command18.Parameters.AddWithValue("@idCard", userid);
                    connection.Open();
                    var reader18 = command18.ExecuteReader();
                    connection.Close();

                    int lotSizeM2 = 0;

                    for (int i = 0; i < thicknessList.Count; i++)
                    {
                        using var command14 = new SqlCommand("DA_GetLastLotHasRemain2", connection) { CommandType = CommandType.StoredProcedure };
                        command14.Parameters.AddWithValue("@idCard", userid);
                        command14.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader14 = command14.ExecuteReader();
                        List<Dictionary<string, object>> lastLotHasRemain = CommonFunction.GetDataFromProcedure(reader14);
                        connection.Close();

                        using var command15 = new SqlCommand("DA_GetLastLotNotHasRemain2", connection) { CommandType = CommandType.StoredProcedure };
                        command15.Parameters.AddWithValue("@idCard", userid);
                        command15.Parameters.AddWithValue("@thickness", thicknessList[i]["thickness"]);
                        connection.Open();
                        var reader15 = command15.ExecuteReader();
                        List<Dictionary<string, object>> lastLotNotHasRemain = CommonFunction.GetDataFromProcedure(reader15);
                        connection.Close();

                        int totalDie = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            totalDie += Convert.ToInt32(lastLotHasRemain[j]["remain"]);
                        }
                        for (int j = 0; j < lastLotNotHasRemain.Count; j++)
                        {
                            totalDie += lastLotNotHasRemain[j]["die_quantity"].ToString() == "" ? 0 : Convert.ToInt32(lastLotNotHasRemain[j]["die_quantity"]);
                        }
                        //End tinh total die

                        //if (thicknessList[i]["thickness"].ToString() == typeThickness)
                        //{
                        //    lotSizeM2 = totalDie;
                        //}

                        //int dieEachAlloc = lotSizeM2 / Convert.ToInt32(thicknessList[i]["total_layer"]) * Convert.ToInt32(thicknessList[i]["die_usage"]);


                        if (thicknessList[i]["thickness"].ToString() == typeThickness)
                        {
                            lotSizeM2 = totalDie / Convert.ToInt32(thicknessList[i]["total_layer"]) / Convert.ToInt32(thicknessList[i]["die_usage"]);
                        }

                        int dieEachAlloc = lotSizeM2 * Convert.ToInt32(thicknessList[i]["die_usage"]);

                        int remain = 0;
                        int indexNotHasRemain = 0;
                        for (int j = 0; j < lastLotHasRemain.Count; j++)
                        {
                            int totalEachAlloc = Convert.ToInt32(lastLotHasRemain[j]["remain"]) + remain;
                            int isFirst = -1;
                            for (int k = indexNotHasRemain; k < lastLotNotHasRemain.Count; k++)
                            {
                                totalEachAlloc += lastLotNotHasRemain[k]["die_quantity"].ToString() == "" ? 0 : Convert.ToInt32(lastLotNotHasRemain[k]["die_quantity"]);

                                using var command16 = new SqlCommand("DA_AddDieAllocLastLot2", connection) { CommandType = CommandType.StoredProcedure };
                                command16.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                command16.Parameters.AddWithValue("@die_alloc", lastLotHasRemain[j]["die_alloc"]);
                                command16.Parameters.AddWithValue("@dieEachAlloc", dieEachAlloc);
                                command16.Parameters.AddWithValue("@lot_size", lotSizeM2);
                                connection.Open();
                                var reader16 = command16.ExecuteReader();
                                connection.Close();

                                if (totalEachAlloc > dieEachAlloc || k == lastLotNotHasRemain.Count - 1)
                                {
                                    isFirst++;
                                    remain = totalEachAlloc - dieEachAlloc;
                                    string waferLotDcc = "";
                                    
                                    using var command17 = new SqlCommand("DA_AddRemainLastLot2", connection) { CommandType = CommandType.StoredProcedure };
                                    command17.Parameters.AddWithValue("@id", lastLotNotHasRemain[k]["id"]);
                                    command17.Parameters.AddWithValue("@remain", remain);
                                    command17.Parameters.AddWithValue("@toOpr", j == lastLotHasRemain.Count - 1 ? "" : lastLotHasRemain[j + 1]["da_opr"]);
                                    command17.Parameters.AddWithValue("@waferLotDcc", waferLotDcc);

                                    connection.Open();
                                    var reader17 = command17.ExecuteReader();
                                    connection.Close();

                                    indexNotHasRemain = k + 1;
                                    break;
                                }
                            }
                        }
                    }
                }

                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Save slice for thickness successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Save slice for thickness successfully",
                    Data = null,
                    size = data.Count
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        [HttpPost]
        public IActionResult ResetSliceForThickness([FromBody] Dictionary<string, object> data)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            string userid = data["userid"].ToString();
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);

                using var command = new SqlCommand("DA_ResetSliceForThickness", connection) { CommandType = CommandType.StoredProcedure };
                command.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader = command.ExecuteReader();
                List<Dictionary<string, object>> rawData = CommonFunction.GetDataFromProcedure(reader);
                connection.Close();

                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Reset slice for thickness successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Reset slice for thickness successfully",
                    Data = null,
                    size = data.Count
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        [HttpGet]
        public IActionResult Export(string userid)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                using var connection = new SqlConnection(CommonFunction.connectionString);

                using var command = new SqlCommand("DA_GetHeader1", connection) { CommandType = CommandType.StoredProcedure };
                command.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader = command.ExecuteReader();
                List<Dictionary<string, object>> header1 = CommonFunction.GetDataFromProcedure(reader);
                connection.Close();

                using var command2 = new SqlCommand("DA_GetHeader2", connection) { CommandType = CommandType.StoredProcedure };
                command2.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader2 = command2.ExecuteReader();
                List<Dictionary<string, object>> header2 = CommonFunction.GetDataFromProcedure(reader2);
                connection.Close();

                using var command3 = new SqlCommand("DA_GetSpecForExcel", connection) { CommandType = CommandType.StoredProcedure };
                command3.Parameters.AddWithValue("@idCard", userid);
                connection.Open();
                var reader3 = command3.ExecuteReader();
                List<Dictionary<string, object>> spec = CommonFunction.GetDataFromProcedure(reader3);
                int daOPR = spec.Count;
                connection.Close();

                //Export excel
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string excelFilePath = Path.Combine(documentsPath, "Loading plan.xlsx");

                string[] sheetNames = { "Sheet1", "Sheet2" };
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage())
                {
                    foreach (string sheetName in sheetNames)
                    {
                        package.Workbook.Worksheets.Add(sheetName);
                    }

                    //Sheet 1
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetNames[0]];

                    // Điền dữ liệu vào sheet
                    worksheet.Cells[2, 2].Value = "Target device";
                    worksheet.Cells[3, 2].Value = "PDL:";
                    worksheet.Cells[4, 2].Value = "MARKING:";
                    worksheet.Cells[5, 2].Value = "B/D:";
                    worksheet.Cells[6, 2].Value = "body size:";
                    worksheet.Cells[7, 2].Value = "PCB:";
                    worksheet.Cells[8, 2].Value = "WAFER NAME: ";
                    worksheet.Cells[9, 2].Value = "MAPPING";
                    worksheet.Cells[10, 2].Value = "PO";
                    worksheet.Cells[11, 2].Value = "Ship to";
                    worksheet.Cells[12, 2].Value = "NPI Flag";
                    worksheet.Cells[13, 2].Value = "Lot Type";
                    worksheet.Cells[14, 2].Value = "Device";
                    worksheet.Cells[15, 2].Value = "Cust Info";

                    // Tô màu xám cho vùng từ A1 đến Z16
                    worksheet.Cells["A1:Z16"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells["A1:Z16"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);

                    // Tô màu cho các header header 1
                    for (int row = 2; row <= 15; row++)
                    {
                        worksheet.Cells[row, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        //worksheet.Cells[row, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        worksheet.Cells[row, 2].Style.Font.Bold = true;

                        // Thêm border cho cell
                        worksheet.Cells[row, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    int colHeader1 = 3;
                    foreach (Dictionary<string, object> colData in header1)
                    {
                        int rowHeader1 = 2;
                        foreach (var keyValuePair in colData) // Duyệt qua từng cặp key-value trong Dictionary
                        {
                            worksheet.Cells[rowHeader1, colHeader1].Value = keyValuePair.Value; // Lấy giá trị của keyValuePair

                            worksheet.Cells[rowHeader1, colHeader1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[rowHeader1, colHeader1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                            worksheet.Cells[rowHeader1, colHeader1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader1, colHeader1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader1, colHeader1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader1, colHeader1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rowHeader1++;
                        }
                        colHeader1++;
                    }

                    //Spec
                    int colSpec = colHeader1 + 1;
                    worksheet.Cells[2, colSpec].Value = "DA OPR";
                    worksheet.Cells[2, colSpec + 1].Value = "ASSM Die alloc";
                    worksheet.Cells[2, colSpec + 2].Value = "BG Thickness";
                    for (int col = colSpec; col <= colSpec + 2; col++)
                    {
                        worksheet.Cells[2, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[2, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        //worksheet.Cells[2, col].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        worksheet.Cells[2, col].Style.Font.Bold = true;

                        // Thêm border cho cell
                        worksheet.Cells[2, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    int rowSpec = 3;
                    foreach (Dictionary<string, object> rowData in spec)
                    {
                        int col = colSpec;
                        foreach (var keyValuePair in rowData) // Duyệt qua từng cặp key-value trong Dictionary
                        {
                            worksheet.Cells[rowSpec, col].Value = keyValuePair.Value; // Lấy giá trị của keyValuePair

                            worksheet.Cells[rowSpec, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[rowSpec, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                            worksheet.Cells[rowSpec, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowSpec, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowSpec, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowSpec, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            col++;
                        }
                        rowSpec++;
                    }

                    //Header 2
                    int colHeader2 = colSpec + 4;
                    worksheet.Cells[2, colHeader2].Value = "";
                    worksheet.Cells[2, colHeader2 + 1].Value = "FG";
                    worksheet.Cells[2, colHeader2 + 2].Value = "PV";
                    worksheet.Cells[2, colHeader2 + 3].Value = "CT";
                    worksheet.Cells[2, colHeader2 + 4].Value = "Line";
                    worksheet.Cells[2, colHeader2 + 5].Value = "Ship to";
                    worksheet.Cells[2, colHeader2 + 6].Value = "STATUS";
                    for (int col = colHeader2; col <= colHeader2 + 6; col++)
                    {
                        worksheet.Cells[2, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[2, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        //worksheet.Cells[2, col].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        worksheet.Cells[2, col].Style.Font.Bold = true;

                        // Thêm border cho cell
                        worksheet.Cells[2, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    int rowHeader2 = 3;
                    foreach (Dictionary<string, object> rowData in header2)
                    {
                        int col = colHeader2;
                        foreach (var keyValuePair in rowData) // Duyệt qua từng cặp key-value trong Dictionary
                        {
                            worksheet.Cells[rowHeader2, col].Value = keyValuePair.Value; // Lấy giá trị của keyValuePair

                            worksheet.Cells[rowHeader2, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[rowHeader2, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                            worksheet.Cells[rowHeader2, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader2, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader2, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[rowHeader2, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            col++;
                        }
                        rowHeader2++;
                    }

                    // Get total thickness and die usage
                    using var command4 = new SqlCommand("DA_GetThickness", connection) { CommandType = CommandType.StoredProcedure };
                    command4.Parameters.AddWithValue("@idCard", userid);
                    connection.Open();
                    var reader4 = command4.ExecuteReader();
                    List<Dictionary<string, object>> thickness = CommonFunction.GetDataFromProcedure(reader4);
                    int totalThickness = thickness.Count;
                    connection.Close();

                    // Get lot size
                    using var command6 = new SqlCommand("DA_GetLotSize", connection) { CommandType = CommandType.StoredProcedure };
                    command6.Parameters.AddWithValue("@idCard", userid);
                    connection.Open();
                    var reader6 = command6.ExecuteReader();
                    List<Dictionary<string, object>> lotSizeObj = CommonFunction.GetDataFromProcedure(reader6);
                    int lotSize = Convert.ToInt32(lotSizeObj[0]["lot_size"]);
                    connection.Close();

                    // Get spec full info
                    using var command7 = new SqlCommand("DA_GetSpec", connection) { CommandType = CommandType.StoredProcedure };
                    command7.Parameters.AddWithValue("@idCard", userid);
                    connection.Open();
                    var reader7 = command7.ExecuteReader();
                    List<Dictionary<string, object>> specFull = CommonFunction.GetDataFromProcedure(reader7);
                    connection.Close();

                    int rowStart = 18;
                    List<int> rowExcel = new List<int>();
                    

                    for(int i = 0; i < specFull.Count; i++)
                    {
                        // GetRawData by die alloc
                        if (i == 0)
                        {
                            using var command5 = new SqlCommand("DA_GetRawDataByDieAlloc", connection) { CommandType = CommandType.StoredProcedure };
                            command5.Parameters.AddWithValue("@idCard", userid);
                            command5.Parameters.AddWithValue("@dieAlloc", specFull[i]["die_alloc"]);
                            connection.Open();
                            var reader5 = command5.ExecuteReader();
                            List<Dictionary<string, object>> rawData = CommonFunction.GetDataFromProcedure(reader5);
                            connection.Close();

                            List<Dictionary<string, object>> requireWaferLot = new List<Dictionary<string, object>>();
                            Dictionary<string, object> lastRequireWaferLot = new Dictionary<string, object>();
                            int fromWafer = 0;
                            for (int j = 0; j < rawData.Count; j++)
                            {
                                requireWaferLot.Add(rawData[j]);
                                if (rawData[j]["remain"].ToString() != "")
                                {
                                    int rowBeforeStart = rowStart - 1;
                                    worksheet.Cells[rowBeforeStart, 1].Value = "M / D";
                                    worksheet.Cells[rowBeforeStart, 2].Value = "ASSY LOT# / DCC";
                                    worksheet.Cells[rowBeforeStart, 3].Value = "Target device";
                                    worksheet.Cells[rowBeforeStart, 4].Value = "Batch#";
                                    worksheet.Cells[rowBeforeStart, 5].Value = "Die Location";
                                    worksheet.Cells[rowBeforeStart, 6].Value = "Wafer Mapping#";
                                    worksheet.Cells[rowBeforeStart, 7].Value = "Wafer pcs";
                                    worksheet.Cells[rowBeforeStart, 8].Value = "AO LOT#/DCC";
                                    worksheet.Cells[rowBeforeStart, 9].Value = "Need Die QTY";
                                    worksheet.Cells[rowBeforeStart, 10].Value = "Loading QTY";
                                    worksheet.Cells[rowBeforeStart, 11].Value = "From wafer";
                                    worksheet.Cells[rowBeforeStart, 12].Value = "Required wafer";
                                    worksheet.Cells[rowBeforeStart, 13].Value = "Remain wafer";
                                    worksheet.Cells[rowBeforeStart, 14].Value = "Loading date";

                                    for (int col = 1; col <= 14; col++)
                                    {
                                        worksheet.Cells[rowBeforeStart, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheet.Cells[rowBeforeStart, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                                        worksheet.Cells[rowBeforeStart, col].Style.Font.Bold = true;
                                    }

                                    //chi in ra lot size dong dau tien
                                    worksheet.Cells[rowStart, 10].Value = lotSize;

                                    var groupedData = requireWaferLot.GroupBy(d => d["wafer_lot"]);
                                    int indexGroupData = 0;
                                    int lengthGroupData = groupedData.Count();
                                    int countNeedDieQty = 0;
                                    foreach (var group in groupedData)
                                    {
                                        indexGroupData++;
                                        string waferLotName = group.Key.ToString();
                                        string waferLotDcc = "";
                                        string batch = "";
                                        string waferMapping = "";
                                        List<string> requireWafer = new List<string>();
                                        int needDieQty = 0;
                                        foreach (var item in group)
                                        {
                                            requireWafer.Add(Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString());
                                            if(item["remain"].ToString() == "" ) needDieQty += Convert.ToInt32(item["slice_die_quantity"]);
                                            waferLotDcc = item["wafer_lot_dcc"].ToString();
                                            batch = item["batch"].ToString();
                                            waferMapping = item["wafer_mapping"].ToString();
                                        }

                                        worksheet.Cells[rowStart, 1].Value = specFull[i]["lot_type"];
                                        worksheet.Cells[rowStart, 2].Value = waferLotDcc;
                                        worksheet.Cells[rowStart, 3].Value = header1[0]["target_device"];
                                        worksheet.Cells[rowStart, 4].Value = batch;
                                        worksheet.Cells[rowStart, 5].Value = specFull[i]["die_alloc"];
                                        worksheet.Cells[rowStart, 6].Value = waferMapping;
                                        worksheet.Cells[rowStart, 7].Value
                                            = lastRequireWaferLot.Count == 0 || lastRequireWaferLot["wafer_lot"].ToString() != waferLotName
                                            ? requireWafer.Count : requireWafer.Count - 1;
                                        worksheet.Cells[rowStart, 8].Value = waferLotName;
                                        if (lengthGroupData > 1)
                                        {
                                            if (indexGroupData == 1)
                                            {
                                                worksheet.Cells[rowStart, 9].Value = needDieQty + Convert.ToInt32(fromWafer);
                                                countNeedDieQty = needDieQty + Convert.ToInt32(fromWafer);
                                            }
                                            else if (indexGroupData != 1 && indexGroupData != lengthGroupData)
                                            {
                                                worksheet.Cells[rowStart, 9].Value = needDieQty;
                                                countNeedDieQty += needDieQty;
                                            }
                                            else
                                            {
                                                worksheet.Cells[rowStart, 9].Value = lotSize * Convert.ToInt32(specFull[i]["die_usage"]) - countNeedDieQty;
                                            }
                                        }
                                        else
                                        {
                                            worksheet.Cells[rowStart, 9].Value = lotSize * Convert.ToInt32(specFull[i]["die_usage"]); //need
                                        }
                                        //worksheet.Cells[rowStart, 10].Value = specFull[i]["lot_type"].ToString() == "M" ? lotSize : "";
                                        worksheet.Cells[rowStart, 11].Value
                                            = lastRequireWaferLot.Count == 0 || lastRequireWaferLot["wafer_lot"].ToString() != waferLotName
                                            ? "" : fromWafer + "ea from #" + (Convert.ToInt32(lastRequireWaferLot["slice_number"]) < 10 ? "0" + lastRequireWaferLot["slice_number"].ToString() : lastRequireWaferLot["slice_number"].ToString());
                                        //worksheet.Cells[rowStart, 12].Value = "#" + requireWafer[0] + "~" + requireWafer[requireWafer.Count - 1];
                                        worksheet.Cells[rowStart, 12].Value = "#" + string.Join(",", requireWafer); ;
                                        worksheet.Cells[rowStart, 13].Value
                                            = lengthGroupData == indexGroupData
                                            ? "#" + requireWafer[requireWafer.Count - 1] + " remain " + rawData[j]["remain"] + "ea to " + specFull[i]["da_opr"] : "";
                                        worksheet.Cells[rowStart, 14].Value = "";

                                        worksheet.Cells[rowStart, 12].Style.Font.Bold = true;
                                        worksheet.Cells[rowStart, 11].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                        worksheet.Cells[rowStart, 13].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                                        rowStart += 1;
                                        if (lengthGroupData == indexGroupData) rowExcel.Add(rowStart);

                                    }
                                    if (Convert.ToInt32(rawData[j]["remain"]) != 0)
                                    {
                                        lastRequireWaferLot = requireWaferLot[requireWaferLot.Count - 1];
                                        requireWaferLot = new List<Dictionary<string, object>>();
                                        requireWaferLot.Add(lastRequireWaferLot);
                                    }
                                    else
                                    {
                                        requireWaferLot = new List<Dictionary<string, object>>();
                                    }
                                    fromWafer = Convert.ToInt32(rawData[j]["remain"]);
                                    rowStart += 20;
                                }
                            }
                        }
                        else
                        {
                            using var command5 = new SqlCommand("DA_GetRawDataByDieAlloc", connection) { CommandType = CommandType.StoredProcedure };
                            command5.Parameters.AddWithValue("@idCard", userid);
                            command5.Parameters.AddWithValue("@dieAlloc", specFull[i]["die_alloc"]);
                            connection.Open();
                            var reader5 = command5.ExecuteReader();
                            List<Dictionary<string, object>> rawData = CommonFunction.GetDataFromProcedure(reader5);
                            connection.Close();

                            List<Dictionary<string, object>> requireWaferLot = new List<Dictionary<string, object>>();
                            Dictionary<string, object> lastRequireWaferLot = new Dictionary<string, object>();
                            int fromWafer = 0;
                            for (int j = 0; j < rawData.Count; j++)
                            {
                                requireWaferLot.Add(rawData[j]);
                                if (rawData[j]["remain"].ToString() != "")
                                {
                                    int indexRemain = 0;

                                    var groupedData = requireWaferLot.GroupBy(d => d["wafer_lot"]);
                                    int indexGroupData = 0;
                                    int lengthGroupData = groupedData.Count();
                                    int countNeedDieQty = 0;
                                    foreach (var group in groupedData)
                                    {
                                        indexGroupData++;
                                        string waferLotName = group.Key.ToString();
                                        string waferLotDcc = "";
                                        string batch = "";
                                        string waferMapping = "";
                                        List<string> requireWafer = new List<string>();
                                        int needDieQty = 0;
                                        bool isFirstItem = true;
                                        foreach (var item in group)
                                        {
                                            requireWafer.Add(Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString());
                                            if (item["remain"].ToString() == "") needDieQty += Convert.ToInt32(item["slice_die_quantity"]);
                                            waferLotDcc = item["wafer_lot_dcc"].ToString();
                                            batch = item["batch"].ToString();
                                            waferMapping = item["wafer_mapping"].ToString();
                                        }

                                        worksheet.Cells[rowExcel[indexRemain], 1].Value = specFull[i]["lot_type"];
                                        worksheet.Cells[rowExcel[indexRemain], 2].Value = waferLotDcc;
                                        worksheet.Cells[rowExcel[indexRemain], 3].Value = header1[0]["target_device"];
                                        worksheet.Cells[rowExcel[indexRemain], 4].Value = batch;
                                        worksheet.Cells[rowExcel[indexRemain], 5].Value = specFull[i]["die_alloc"];
                                        worksheet.Cells[rowExcel[indexRemain], 6].Value = waferMapping;
                                        worksheet.Cells[rowExcel[indexRemain], 7].Value
                                            = lastRequireWaferLot.Count == 0 || lastRequireWaferLot["wafer_lot"].ToString() != waferLotName
                                            ? requireWafer.Count : requireWafer.Count - 1;
                                        worksheet.Cells[rowExcel[indexRemain], 8].Value = waferLotName;
                                        if (lengthGroupData > 1)
                                        {
                                            if (indexGroupData == 1)
                                            {
                                                worksheet.Cells[rowExcel[indexRemain], 9].Value = needDieQty + Convert.ToInt32(fromWafer);
                                                countNeedDieQty = needDieQty + Convert.ToInt32(fromWafer);
                                            }
                                            else if (indexGroupData != 1 && indexGroupData != lengthGroupData)
                                            {
                                                worksheet.Cells[rowExcel[indexRemain], 9].Value = needDieQty;
                                                countNeedDieQty += needDieQty;
                                            }
                                            else
                                            {
                                                worksheet.Cells[rowExcel[indexRemain], 9].Value = lotSize * Convert.ToInt32(specFull[i]["die_usage"]) - countNeedDieQty;
                                            }
                                        }
                                        else
                                        {
                                            worksheet.Cells[rowExcel[indexRemain], 9].Value = lotSize * Convert.ToInt32(specFull[i]["die_usage"]); //need
                                        }
                                        //worksheet.Cells[rowExcel[indexRemain], 10].Value = specFull[i]["lot_type"].ToString() == "M" ? lotSize : "";
                                        worksheet.Cells[rowExcel[indexRemain], 11].Value
                                            = lastRequireWaferLot.Count == 0 || lastRequireWaferLot["wafer_lot"].ToString() != waferLotName
                                            ? "" : fromWafer + "ea from #" + (Convert.ToInt32(lastRequireWaferLot["slice_number"]) < 10 ? "0" + lastRequireWaferLot["slice_number"].ToString() : lastRequireWaferLot["slice_number"].ToString());
                                        //worksheet.Cells[rowExcel[indexRemain], 12].Value = "#" + requireWafer[0] + "~" + requireWafer[requireWafer.Count - 1];
                                        worksheet.Cells[rowExcel[indexRemain], 12].Value = "#" + string.Join(",", requireWafer);
                                        worksheet.Cells[rowExcel[indexRemain], 13].Value
                                            = lengthGroupData == indexGroupData
                                            ? "#" + requireWafer[requireWafer.Count - 1] + " remain " + rawData[j]["remain"] + "ea to " + specFull[i]["da_opr"] : "";
                                        worksheet.Cells[rowExcel[indexRemain], 14].Value = "";

                                        worksheet.Cells[rowExcel[indexRemain], 12].Style.Font.Bold = true;
                                        worksheet.Cells[rowExcel[indexRemain], 11].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                        worksheet.Cells[rowExcel[indexRemain], 13].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                                        rowExcel[indexRemain] += 1;
                                        if (lengthGroupData == indexGroupData)
                                        {
                                            rowExcel.Add(rowExcel[indexRemain]);
                                            rowExcel.RemoveAt(indexRemain);
                                        }
                                    }
                                    if (Convert.ToInt32(rawData[j]["remain"]) != 0)
                                    {
                                        lastRequireWaferLot = requireWaferLot[requireWaferLot.Count - 1];
                                        requireWaferLot = new List<Dictionary<string, object>>();
                                        requireWaferLot.Add(lastRequireWaferLot);
                                    }
                                    else
                                    {
                                        requireWaferLot = new List<Dictionary<string, object>>();
                                    }
                                    fromWafer = Convert.ToInt32(rawData[j]["remain"]);
                                }
                            }
                        }
                    }

                    //Tinh toan nhung die con thua con lai
                    //Get raw data last lot 1
                    using var command8 = new SqlCommand("DA_GetRawLastLot", connection) { CommandType = CommandType.StoredProcedure };
                    command8.Parameters.AddWithValue("@idCard", userid);
                    //command8.Parameters.AddWithValue("@thickness", thickness[i]["thickness"]);
                    connection.Open();
                    var reader8 = command8.ExecuteReader();
                    List<Dictionary<string, object>> rawDataLastLot = CommonFunction.GetDataFromProcedure(reader8);
                    connection.Close();

                    //Get raw data last lot 2
                    using var command9 = new SqlCommand("DA_GetRawLastLot2", connection) { CommandType = CommandType.StoredProcedure };
                    command9.Parameters.AddWithValue("@idCard", userid);
                    connection.Open();
                    var reader9 = command9.ExecuteReader();
                    List<Dictionary<string, object>> rawDataLastLot2 = CommonFunction.GetDataFromProcedure(reader9);
                    connection.Close();

                    if(rawDataLastLot2.Count > 0)
                    {
                        //Get loading qty
                        using var command10 = new SqlCommand("DA_GetLoadingQty", connection) { CommandType = CommandType.StoredProcedure };
                        command10.Parameters.AddWithValue("@idCard", userid);
                        connection.Open();
                        var reader10 = command10.ExecuteReader();
                        List<Dictionary<string, object>> loadingQty = CommonFunction.GetDataFromProcedure(reader10);
                        connection.Close();

                        int startRowLastLot = rowExcel[rowExcel.Count - 1];

                        worksheet.Cells[startRowLastLot, 1].Value = "M / D";
                        worksheet.Cells[startRowLastLot, 2].Value = "ASSY LOT# / DCC";
                        worksheet.Cells[startRowLastLot, 3].Value = "Target device";
                        worksheet.Cells[startRowLastLot, 4].Value = "Batch#";
                        worksheet.Cells[startRowLastLot, 5].Value = "Die Location";
                        worksheet.Cells[startRowLastLot, 6].Value = "Wafer Mapping#";
                        worksheet.Cells[startRowLastLot, 7].Value = "Wafer pcs";
                        worksheet.Cells[startRowLastLot, 8].Value = "AO LOT#/DCC";
                        worksheet.Cells[startRowLastLot, 9].Value = "Need Die QTY";
                        worksheet.Cells[startRowLastLot, 10].Value = "Loading QTY";
                        worksheet.Cells[startRowLastLot, 11].Value = "From wafer";
                        worksheet.Cells[startRowLastLot, 12].Value = "Required wafer";
                        worksheet.Cells[startRowLastLot, 13].Value = "Remain wafer";
                        worksheet.Cells[startRowLastLot, 14].Value = "Loading date";
                        for (int col = 1; col <= 14; col++)
                        {
                            worksheet.Cells[startRowLastLot, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[startRowLastLot, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            worksheet.Cells[startRowLastLot, col].Style.Font.Bold = true;
                        }

                        //startRowLastLot++;
                        List<int> rowsExcelLastLot = new List<int>();
                        for (int i = 0; i < thickness.Count; i++)
                        {
                            startRowLastLot += 20;
                            rowsExcelLastLot.Add(startRowLastLot);
                        }
                        //Them loading dong dau tien
                        worksheet.Cells[rowsExcelLastLot[0], 10].Value = Convert.ToInt32(loadingQty[0]["loading_qty"]);

                        var groupedLastLot = rawDataLastLot.GroupBy(d => d["thickness"]);
                        int indexThicknessList = -1;
                        int startRowLastLot2 = 200;
                        foreach (var thicknessList in groupedLastLot)
                        {
                            indexThicknessList++;
                            int stRow = rowsExcelLastLot[indexThicknessList];

                            var groupedDieAlloc = thicknessList.GroupBy(d => d["die_alloc"]);
                            int indexDieAlloc = -1;
                            int remainInThickness = 0;
                            string sliceNumberInThickness = "";
                            foreach (var dieAllocList in groupedDieAlloc)
                            {
                                indexDieAlloc++;
                                string dieAllocName = dieAllocList.Key.ToString();
                                var groupedWaferLot = dieAllocList.GroupBy(d => d["wafer_lot"]);

                                foreach (var waferLot in groupedWaferLot)
                                {
                                    string lotType = "";
                                    string lotDcc = "";
                                    string batch = "";
                                    string mapping = "";
                                    int waferPCS = 0;
                                    string aoLot = "";
                                    int loading = 0;
                                    string fromWafer = "";
                                    List<string> requireWafer = new List<string>();
                                    int remain = 0;
                                    string toOpr = "";
                                    int needDie = 0;
                                    foreach (var item in waferLot)
                                    {
                                        lotType = item["lot_type"].ToString();
                                        lotDcc = item["wafer_lot_dcc"].ToString();
                                        batch = item["batch"].ToString();
                                        mapping = item["wafer_mapping"].ToString();
                                        aoLot = item["wafer_lot"].ToString();
                                        if (Convert.ToInt32(item["isFullLotRemain"]) == 1)
                                        {
                                            needDie = Convert.ToInt32(item["remain"]);
                                        }
                                        else
                                        {
                                            needDie += Convert.ToInt32(item["die_quantity"]) - (item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]));
                                        }

                                        loading = item["dieEachAlloc"].ToString() == "" ? 0 : Convert.ToInt32(item["dieEachAlloc"]);
                                        requireWafer.Add((Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString()));
                                        if (Convert.ToInt32(item["isFullLotRemain"]) == 1)
                                        {
                                            fromWafer = item["remain"] + "ea from #" + (Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString());
                                            waferPCS++;
                                            if (indexDieAlloc != 0 && rawDataLastLot2.Count == 0)
                                            {
                                                fromWafer += "; " + remainInThickness + "ea from #" + sliceNumberInThickness;
                                                requireWafer.Add(sliceNumberInThickness);
                                                needDie += remainInThickness;
                                                waferPCS++;
                                            }
                                        }
                                        if (Convert.ToInt32(item["isFullLotRemain"]) == 0)
                                        {
                                            remain = item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]);
                                            remainInThickness = item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]);
                                        }
                                        toOpr = item["to_opr"].ToString();

                                    }
                                    worksheet.Cells[stRow, 1].Value = lotType;
                                    worksheet.Cells[stRow, 2].Value = lotDcc;
                                    worksheet.Cells[stRow, 3].Value = header1[0]["target_device"];
                                    worksheet.Cells[stRow, 4].Value = batch;
                                    worksheet.Cells[stRow, 5].Value = dieAllocName;
                                    worksheet.Cells[stRow, 6].Value = mapping;
                                    worksheet.Cells[stRow, 7].Value = requireWafer.Count - waferPCS;
                                    worksheet.Cells[stRow, 8].Value = aoLot;
                                    worksheet.Cells[stRow, 9].Value = needDie;
                                    //worksheet.Cells[stRow, 10].Value = loading;
                                    worksheet.Cells[stRow, 11].Value = fromWafer;
                                    //worksheet.Cells[stRow, 12].Value = "#"+ requireWafer[0] +"~"+ requireWafer[requireWafer.Count -1];
                                    worksheet.Cells[stRow, 12].Value = "#" + string.Join(",", requireWafer);
                                    worksheet.Cells[stRow, 13].Value = remain == 0 ? "" : "#" + requireWafer[requireWafer.Count - 1] + " remain " + remain + "ea to " + toOpr;
                                    worksheet.Cells[stRow, 14].Value = "";

                                    worksheet.Cells[stRow, 12].Style.Font.Bold = true;
                                    worksheet.Cells[stRow, 11].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                    worksheet.Cells[stRow, 13].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                                    sliceNumberInThickness = requireWafer[requireWafer.Count - 1];
                                    stRow++;
                                }
                                stRow += 50;
                                startRowLastLot2 = stRow + 1;
                            }
                        }

                        // Last lot 2
                        //Get loading qty 2
                        using var command11 = new SqlCommand("DA_GetLoadingQty2", connection) { CommandType = CommandType.StoredProcedure };
                        command11.Parameters.AddWithValue("@idCard", userid);
                        connection.Open();
                        var reader11 = command11.ExecuteReader();
                        List<Dictionary<string, object>> loadingQty2 = CommonFunction.GetDataFromProcedure(reader11);
                        connection.Close();

                        LastLot2(loadingQty2, startRowLastLot2, thickness, rawDataLastLot2, header1[0]["target_device"].ToString(), worksheet);

                    }
                    else
                    {
                        //Get loading qty
                        using var command10 = new SqlCommand("DA_GetLoadingQty", connection) { CommandType = CommandType.StoredProcedure };
                        command10.Parameters.AddWithValue("@idCard", userid);
                        connection.Open();
                        var reader10 = command10.ExecuteReader();
                        List<Dictionary<string, object>> loadingQty = CommonFunction.GetDataFromProcedure(reader10);
                        connection.Close();

                        int startRowLastLot = rowExcel[rowExcel.Count - 1];

                        LastLot2(loadingQty, startRowLastLot, thickness, rawDataLastLot, header1[0]["target_device"].ToString(), worksheet);

                    }

                    // Đóng băng các dòng từ 1 đến 17
                    worksheet.View.FreezePanes(18, 1);

                    int lastRow = worksheet.Dimension.End.Row;
                    if (lastRow > 18) // Kiểm tra xem có dòng nào sau dòng 18 hay không
                    {
                        for (int i = lastRow; i > 18; i--) // Duyệt từ dòng cuối cùng lên dòng 19
                        {
                            if (worksheet.Cells[i, 1].Value == null || worksheet.Cells[i, 1].Value.ToString().Trim() == "")
                            {
                                worksheet.DeleteRow(i);
                            }
                        }
                    }

                    int lastRow2 = worksheet.Dimension.End.Row;
                    List<Dictionary<string, object>> aoLotList = new List<Dictionary<string, object>>();
                    for (int i = 18; i <= lastRow2; i++) 
                    {
                        Dictionary<string, object> rowLot = new Dictionary<string, object>();
                        if (worksheet.Cells[i, 8].Value.ToString() == "AO LOT#/DCC") continue;
                        rowLot.Add("aoLot", worksheet.Cells[i, 8].Value);
                        rowLot.Add("row", i);
                        aoLotList.Add(rowLot);
                    }
                    var groupAoLot = aoLotList.GroupBy(d => d["aoLot"]);
                    foreach( var aoLot in groupAoLot)
                    {
                        int dcc = 0;

                        foreach (var item in aoLot)
                        {
                            dcc++;
                            int row = Convert.ToInt32(item["row"]);
                            string dccString = dcc < 10 ? "0" + dcc : dcc.ToString();
                            worksheet.Cells[row, 2].Value = aoLot.Key.ToString() + "/" + dccString;
                        }
                    }
                    for(int row = 1; row<= worksheet.Dimension.End.Row; row++)
                    {
                        for(int col = 1; col<= 26; col++)
                        {
                            worksheet.Cells[row, col].Style.Font.Name = "Tahoma";
                            worksheet.Cells[row, col].Style.Font.Size = 9;
                            if(row >= 17 && col <=8) worksheet.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            if(row >= 17 && col <=14) BorderCell(row, col, worksheet);
                        }
                    }

                    // Căn chỉnh độ rộng cột tự động
                    worksheet.Cells.AutoFitColumns();

                    //Sheet 2
                    ExcelWorksheet worksheet2 = package.Workbook.Worksheets[sheetNames[1]];
                    DrawSheet2(worksheet2, userid);
                    
                    // Lưu file Excel vào MemoryStream
                    using (MemoryStream ms = new MemoryStream())
                    {
                        package.SaveAs(ms);
                        byte[] fileBytes = ms.ToArray();
                        string base64String = Convert.ToBase64String(fileBytes);
                        CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Export successfully", CommonFunction.SUCCESS, functionName);
                        return Ok(new { data = base64String });
                    }
                }
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<String, Object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        public void LastLot2(List<Dictionary<string, object>>loadingQty2
            , int startRowLastLot2
            , List<Dictionary<string, object>> thickness
            , List<Dictionary<string, object>> rawDataLastLot2
            , string targetDevice
            , ExcelWorksheet worksheet)
        {
            worksheet.Cells[startRowLastLot2, 1].Value = "M / D";
            worksheet.Cells[startRowLastLot2, 2].Value = "ASSY LOT# / DCC";
            worksheet.Cells[startRowLastLot2, 3].Value = "Target device";
            worksheet.Cells[startRowLastLot2, 4].Value = "Batch#";
            worksheet.Cells[startRowLastLot2, 5].Value = "Die Location";
            worksheet.Cells[startRowLastLot2, 6].Value = "Wafer Mapping#";
            worksheet.Cells[startRowLastLot2, 7].Value = "Wafer pcs";
            worksheet.Cells[startRowLastLot2, 8].Value = "AO LOT#/DCC";
            worksheet.Cells[startRowLastLot2, 9].Value = "Need Die QTY";
            worksheet.Cells[startRowLastLot2, 10].Value = "Loading QTY";
            worksheet.Cells[startRowLastLot2, 11].Value = "From wafer";
            worksheet.Cells[startRowLastLot2, 12].Value = "Required wafer";
            worksheet.Cells[startRowLastLot2, 13].Value = "Remain wafer";
            worksheet.Cells[startRowLastLot2, 14].Value = "Loading date";
            for (int col = 1; col <= 14; col++)
            {
                worksheet.Cells[startRowLastLot2, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[startRowLastLot2, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[startRowLastLot2, col].Style.Font.Bold = true;
            }

            //startRowLastLot++;
            List<int> rowsExcelLastLot2 = new List<int>();
            for (int i = 0; i < thickness.Count; i++)
            {
                startRowLastLot2 += 20;
                rowsExcelLastLot2.Add(startRowLastLot2);
            }
            worksheet.Cells[rowsExcelLastLot2[0], 10].Value = Convert.ToInt32(loadingQty2[0]["loading_qty"]);

            var groupedLastLot2 = rawDataLastLot2.GroupBy(d => d["thickness"]);
            int indexThicknessList2 = -1;
            foreach (var thicknessList in groupedLastLot2)
            {
                indexThicknessList2++;
                int stRow = rowsExcelLastLot2[indexThicknessList2];

                var groupedDieAlloc = thicknessList.GroupBy(d => d["die_alloc"]);
                int indexDieAlloc = -1;
                int remainInThickness = 0;
                string sliceNumberInThickness = "";
                string aoLotInThickness = "";
                foreach (var dieAllocList in groupedDieAlloc)
                {
                    indexDieAlloc++;
                    string dieAllocName = dieAllocList.Key.ToString();
                    var groupedWaferLot = dieAllocList.GroupBy(d => d["wafer_lot"]);
                    foreach (var waferLot in groupedWaferLot)
                    {
                        string lotType = "";
                        string lotDcc = "";
                        string batch = "";
                        string mapping = "";
                        int waferPCS = 0;
                        string aoLot = "";
                        int loading = 0;
                        string fromWafer = "";
                        List<string> requireWafer = new List<string>();
                        int remain = 0;
                        string toOpr = "";
                        int needDie = 0;
                        int indexWaferLot = -1;
                        foreach (var item in waferLot)
                        {
                            indexWaferLot++;
                            lotType = item["lot_type"].ToString();
                            lotDcc = item["wafer_lot_dcc"].ToString();
                            batch = item["batch"].ToString();
                            mapping = item["wafer_mapping"].ToString();
                            aoLot = item["wafer_lot"].ToString();
                            if (Convert.ToInt32(item["isFullLotRemain"]) == 1)
                            {
                                needDie = Convert.ToInt32(item["remain"]);
                            }
                            else
                            {
                                needDie += Convert.ToInt32(item["die_quantity"]) - (item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]));
                            }

                            loading = item["dieEachAlloc"].ToString() == "" ? 0 : Convert.ToInt32(item["dieEachAlloc"]);
                            requireWafer.Add((Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString()));

                            if (Convert.ToInt32(item["isFullLotRemain"]) == 1)
                            {
                                fromWafer = item["remain"] + "ea from #" + (Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString());
                                waferPCS++;
                                if (indexDieAlloc != 0 && item["wafer_lot"].ToString() == aoLotInThickness)
                                {
                                    fromWafer += "; " + remainInThickness + "ea from #" + sliceNumberInThickness;
                                    requireWafer.Add(sliceNumberInThickness);
                                    needDie += remainInThickness;
                                    waferPCS++;
                                }
                            }
                            else
                            {
                                if (indexDieAlloc != 0 && indexWaferLot == 0 && item["wafer_lot"].ToString() == aoLotInThickness)
                                {
                                    fromWafer = remainInThickness + "ea from #" + sliceNumberInThickness;
                                    requireWafer.Add(sliceNumberInThickness);
                                    needDie += remainInThickness;
                                    waferPCS++;
                                }
                            }
                            //requireWafer.Add((Convert.ToInt32(item["slice_number"]) < 10 ? "0" + item["slice_number"].ToString() : item["slice_number"].ToString()));
                            if (Convert.ToInt32(item["isFullLotRemain"]) == 0)
                            {
                                remain = item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]);
                                remainInThickness = item["remain"].ToString() == "" ? 0 : Convert.ToInt32(item["remain"]);
                                sliceNumberInThickness = requireWafer[requireWafer.Count - 1];
                                aoLotInThickness = aoLot;
                            }

                            toOpr = item["to_opr"].ToString();

                        }
                        worksheet.Cells[stRow, 1].Value = lotType;
                        worksheet.Cells[stRow, 2].Value = lotDcc;
                        worksheet.Cells[stRow, 3].Value = targetDevice;
                        worksheet.Cells[stRow, 4].Value = batch;
                        worksheet.Cells[stRow, 5].Value = dieAllocName;
                        worksheet.Cells[stRow, 6].Value = mapping;
                        worksheet.Cells[stRow, 7].Value = requireWafer.Count - waferPCS;
                        worksheet.Cells[stRow, 8].Value = aoLot;
                        worksheet.Cells[stRow, 9].Value = needDie;
                        //worksheet.Cells[stRow, 10].Value = loading;
                        worksheet.Cells[stRow, 11].Value = fromWafer;
                        //worksheet.Cells[stRow, 12].Value = "#"+ requireWafer[0] +"~"+ requireWafer[requireWafer.Count -1];
                        worksheet.Cells[stRow, 12].Value = "#" + string.Join(",", requireWafer);
                        worksheet.Cells[stRow, 13].Value = remain == 0 ? "" : "#" + requireWafer[requireWafer.Count - 1] + " remain " + remain + "ea to " + toOpr;
                        worksheet.Cells[stRow, 14].Value = "";

                        worksheet.Cells[stRow, 12].Style.Font.Bold = true;
                        worksheet.Cells[stRow, 11].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                        worksheet.Cells[stRow, 13].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                        stRow++;
                    }
                    stRow += 50;
                }
            }
        }
        public void DrawSheet2(ExcelWorksheet worksheet2, string userid)
        {
            worksheet2.Cells[1, 1].Value = "Received Date";
            worksheet2.Cells[1, 2].Value = "Source Device";
            worksheet2.Cells[1, 3].Value = "Key No";
            worksheet2.Cells[1, 4].Value = "Wafer Lot#";

            using var connection = new SqlConnection(CommonFunction.connectionString);
            // Get spec full info
            using var command7 = new SqlCommand("DA_GetSpec", connection) { CommandType = CommandType.StoredProcedure };
            command7.Parameters.AddWithValue("@idCard", userid);
            connection.Open();
            var reader7 = command7.ExecuteReader();
            List<Dictionary<string, object>> specFull = CommonFunction.GetDataFromProcedure(reader7);
            connection.Close();

            var thicknessGroup = specFull.GroupBy(d => d["thickness"]);
            int indexThickss = 5;
            List<string> thicknessList = new List<string>();
            foreach (var thickss in thicknessGroup)
            {
                string thicknessName = thickss.Key.ToString();
                thicknessList.Add(thicknessName);
                string dieAlloc = "";
                int index = -1;
                foreach (var item in thickss)
                {
                    index++;
                    dieAlloc += item["die_alloc"].ToString();
                    if (index < (thickss.Count() - 1))
                    {
                        dieAlloc += ",";
                    }
                }
                string titleThickness = dieAlloc + " (" + thicknessName + ")";
                worksheet2.Cells[1, indexThickss].Value = titleThickness;
                indexThickss++;
            }
            worksheet2.Cells[1, indexThickss].Value = "REMAIN WAFER ID";
            worksheet2.Cells[1, indexThickss + 1].Value = "REMAINING QTY";
            worksheet2.Cells[1, indexThickss + 2].Value = "DETAILS";
            for (int col = 1; col <= indexThickss + 2; col++)
            {
                worksheet2.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet2.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                //worksheet2.Cells[1, col].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet2.Cells[1, col].Style.Font.Bold = true;

                BorderCell(1, col, worksheet2);
            }

            using var command1 = new SqlCommand("DA_GetWaferLot", connection) { CommandType = CommandType.StoredProcedure };
            command1.Parameters.AddWithValue("@idCard", userid);
            connection.Open();
            var reader1 = command1.ExecuteReader();
            List<Dictionary<string, object>> waferLot = CommonFunction.GetDataFromProcedure(reader1);
            connection.Close();
            for(int i = 0; i < waferLot.Count; i++)
            {
                worksheet2.Cells[i + 2, 1].Value = waferLot[i]["receive_date"];
                worksheet2.Cells[i + 2, 2].Value = waferLot[i]["source_device"];
                worksheet2.Cells[i + 2, 3].Value = waferLot[i]["key_no"];
                worksheet2.Cells[i + 2, 4].Value = waferLot[i]["customer_lot"];

                BorderCell(i + 2, 1, worksheet2);
                BorderCell(i + 2, 2, worksheet2);
                BorderCell(i + 2, 3, worksheet2);
                BorderCell(i + 2, 4, worksheet2);

                int indThickss = 5;
                
                List<Dictionary<string, object>> remainList = new List<Dictionary<string, object>>();
                for (int j = 0; j < thicknessList.Count; j++) {
                    using var command2 = new SqlCommand("DA_GetSliceSheet2", connection) { CommandType = CommandType.StoredProcedure };
                    command2.Parameters.AddWithValue("@idCard", userid);
                    command2.Parameters.AddWithValue("@waferLot", waferLot[i]["customer_lot"]);
                    command2.Parameters.AddWithValue("@thickness", thicknessList[j]);
                    connection.Open();
                    var reader2 = command2.ExecuteReader();
                    List<Dictionary<string, object>> sliceNumber = CommonFunction.GetDataFromProcedure(reader2);
                    connection.Close();

                    worksheet2.Cells[i + 2, indThickss].Value = "#" + sliceNumber[0]["slice_number"];
                    worksheet2.Cells[i + 2, indThickss].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 2, indThickss].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 2, indThickss].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet2.Cells[i + 2, indThickss].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    indThickss++;

                    using var command3 = new SqlCommand("DA_GetLastRemainByThickness", connection) { CommandType = CommandType.StoredProcedure };
                    command3.Parameters.AddWithValue("@idCard", userid);
                    command3.Parameters.AddWithValue("@thickness", thicknessList[j]);
                    connection.Open();
                    var reader3 = command3.ExecuteReader();
                    List<Dictionary<string, object>> lastRemain = CommonFunction.GetDataFromProcedure(reader3);
                    connection.Close();

                    Dictionary<string, object> remain = new Dictionary<string, object>();
                    if(waferLot[i]["customer_lot"].ToString() == lastRemain[0]["wafer_lot"].ToString())
                    {
                        remain.Add("slice_number", lastRemain[0]["slice_number"]);
                        remain.Add("remain", lastRemain[0]["remain"]);
                        remainList.Add(remain);
                    }
                }
                string lastSliceNumber = "";
                int lastRemainQty = 0;
                string detail = "";
                for (int j = 0; j < remainList.Count; j++) {
                    detail += remainList[j]["slice_number"].ToString() + ": " + remainList[j]["remain"] + "ea";
                    lastSliceNumber += remainList[j]["slice_number"].ToString();
                    if (j < remainList.Count - 1) {
                        lastSliceNumber += ",";
                        detail += ",";
                    }
                    lastRemainQty += Convert.ToInt32(remainList[j]["remain"]);
                }
                worksheet2.Cells[i + 2, indThickss].Value = "#"+ lastSliceNumber;
                worksheet2.Cells[i + 2, indThickss+1].Value = lastRemainQty;
                worksheet2.Cells[i + 2, indThickss+2].Value = "#" + detail;

                BorderCell(i + 2, indThickss, worksheet2);
                BorderCell(i + 2, indThickss+1, worksheet2);
                BorderCell(i + 2, indThickss+2, worksheet2);
            }
            
            worksheet2.Cells.AutoFitColumns();
        }

        public void BorderCell(int row, int col, ExcelWorksheet worksheet)
        {
            worksheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
        [HttpPost]
        public IActionResult UploadFile(IFormFile file, string userid)
        {
            string functionName = ControllerContext.ActionDescriptor.ControllerName + "/" + System.Reflection.MethodBase.GetCurrentMethod().Name;

            if (file == null || file.Length == 0)
            {
                return BadRequest("Please upload file.");
            }

            if (!file.FileName.EndsWith(".xlsx") && !file.FileName.EndsWith(".xls"))
            {
                return BadRequest("File type is not supported. Please upload file .xls, .xlsx");
            }
            try
            {
                // Sử dụng License Context miễn phí
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //Delete data before insert
                using var connection = new SqlConnection(CommonFunction.connectionString);
                using var command = new SqlCommand("DA_DeleteRawData", connection) { CommandType = CommandType.StoredProcedure };

                command.Parameters.AddWithValue("@idCard", userid);

                connection.Open();
                var reader = command.ExecuteReader();
                connection.Close();

                using (var stream = file.OpenReadStream())
                using (var package = new ExcelPackage(stream))
                {
                    string waferName = "";
                    for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                    {
                        if (i == 0) waferName = UploadSpec(package.Workbook.Worksheets[i], userid); 
                        if (i == 1) UploadRawFile(package.Workbook.Worksheets[i], userid, waferName);
                        if (i == 2) UploadWaferLot(package.Workbook.Worksheets[i], userid);
                        if (i == 3) UploadHeader1(package.Workbook.Worksheets[i], userid);
                        if (i == 4) UploadHeader2(package.Workbook.Worksheets[i], userid);
                    }
                }
                
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, "Upload file successfully", CommonFunction.SUCCESS, functionName);
                var response = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.SUCCESS,
                    Message = "Upload file successfully",
                    Data = null,
                    size = 0
                };
                return Ok(response);
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (ví dụ: log lỗi, trả về phản hồi lỗi)
                CommonFunction.LogInfo(CommonFunction.connectionString, userid, ex.Message, CommonFunction.ERROR, functionName);
                var errorResponse = new CommonResponse<Dictionary<string, object>>
                {
                    StatusCode = CommonFunction.ERROR,
                    Message = ex.Message,
                    Data = null,
                    size = 0
                };
                return StatusCode(500, errorResponse);
            }
        }
        private string UploadSpec(ExcelWorksheet worksheet, string userid)
        {
            string waferName = "";
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                if (row == 1) continue;
                Dictionary<string, string> rowData = new Dictionary<string, string>();
                rowData.Add("id_card", userid);

                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    // Lấy giá trị của cell
                    var cellValue = worksheet.Cells[row, col].Value;
                    switch (col)
                    {
                        case 1:
                            rowData.Add("da_opr", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 2:
                            rowData.Add("die_alloc", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 3:
                            rowData.Add("thickness", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 4:
                            rowData.Add("lot_type", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 5:
                            rowData.Add("die_layout", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 6:
                            rowData.Add("wafer_name", cellValue == null ? "" : cellValue.ToString());
                            waferName = cellValue == null ? "" : cellValue.ToString();
                            continue;
                        case 7:
                            rowData.Add("wafer_type", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 8:
                            rowData.Add("die_usage", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 9:
                            rowData.Add("share_wafer", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 10:
                            rowData.Add("number_of_layer", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 11:
                            rowData.Add("lot_size", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 12:
                            rowData.Add("lot_combine_rule", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        default:
                            continue;
                    }
                }
                dataList.Add(rowData);
            }
            using var connection = new SqlConnection(CommonFunction.connectionString);
            connection.Open();
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_spec";

                // Map các cột từ listRaw vào bảng đích
                bulkCopy.ColumnMappings.Add("da_opr", "da_opr");
                bulkCopy.ColumnMappings.Add("die_alloc", "die_alloc");
                bulkCopy.ColumnMappings.Add("thickness", "thickness");
                bulkCopy.ColumnMappings.Add("lot_type", "lot_type");
                bulkCopy.ColumnMappings.Add("die_layout", "die_layout");
                bulkCopy.ColumnMappings.Add("wafer_name", "wafer_name");
                bulkCopy.ColumnMappings.Add("wafer_type", "wafer_type");
                bulkCopy.ColumnMappings.Add("die_usage", "die_usage");
                bulkCopy.ColumnMappings.Add("share_wafer", "share_wafer");
                bulkCopy.ColumnMappings.Add("number_of_layer", "number_of_layer");
                bulkCopy.ColumnMappings.Add("lot_size", "lot_size");
                bulkCopy.ColumnMappings.Add("lot_combine_rule", "lot_combine_rule");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                // Tạo DataTable từ listRaw
                var dataTable = new DataTable();
                dataTable.Columns.Add("da_opr", typeof(string));
                dataTable.Columns.Add("die_alloc", typeof(string));
                dataTable.Columns.Add("thickness", typeof(string));
                dataTable.Columns.Add("lot_type", typeof(string));
                dataTable.Columns.Add("die_layout", typeof(string));
                dataTable.Columns.Add("wafer_name", typeof(string));
                dataTable.Columns.Add("wafer_type", typeof(string));
                dataTable.Columns.Add("die_usage", typeof(string));
                dataTable.Columns.Add("share_wafer", typeof(string));
                dataTable.Columns.Add("number_of_layer", typeof(string));
                dataTable.Columns.Add("lot_size", typeof(string));
                dataTable.Columns.Add("lot_combine_rule", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in dataList)
                {
                    var row = dataTable.NewRow();
                    row["da_opr"] = item["da_opr"];
                    row["die_alloc"] = item["die_alloc"];
                    row["thickness"] = item["thickness"];
                    row["lot_type"] = item["lot_type"];
                    row["die_layout"] = item["die_layout"];
                    row["wafer_name"] = item["wafer_name"];
                    row["wafer_type"] = item["wafer_type"];
                    row["die_usage"] = item["die_usage"];
                    row["share_wafer"] = item["share_wafer"];
                    row["number_of_layer"] = item["number_of_layer"];
                    row["lot_size"] = item["lot_size"];
                    row["lot_combine_rule"] = item["lot_combine_rule"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                // Sử dụng SqlBulkCopy để lưu dữ liệu vào cơ sở dữ liệu
                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }
            return waferName;
        }
        private void UploadRawFile(ExcelWorksheet worksheet, string userid, string waferName)
        {
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> sliceList = new List<Dictionary<string, string>>();
            string waferLot = "";
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                //if (row == 1) continue;
                Dictionary<string, string> rowData = new Dictionary<string, string>();
                Dictionary<string, string> sliceData = new Dictionary<string, string>();
                rowData.Add("id_card", userid);
                if (worksheet.Cells[row, 8].Value.ToString() != waferName) continue;
                for (int col = worksheet.Dimension.Start.Column; col <= 139; col++)
                {
                    // Lấy giá trị của cell
                    var cellValue = worksheet.Cells[row, col].Value;
                    switch (col)
                    {
                        case 5:
                            rowData.Add("key_no", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 6:
                            waferLot = cellValue == null ? "" : cellValue.ToString();
                            rowData.Add("wafer_lot", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 7:
                            rowData.Add("mapping", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 8:
                            rowData.Add("source_device", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 9:
                            rowData.Add("wafer_quantity", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 10:
                            rowData.Add("die_quantity", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        default:
                            if (col >= 16)
                            {
                                switch (col % 5)
                                {
                                    case 1:
                                        sliceData = new Dictionary<string, string>();
                                        sliceData.Add("wafer_lot", waferLot);
                                        sliceData.Add("id_card", userid);
                                        sliceData.Add("thickness", "");
                                        sliceData.Add("backgroundColor", "");
                                        sliceData.Add("die_alloc", "");
                                        sliceData.Add("wafer_lot_dcc", "");
                                        sliceData.Add("remain", "");
                                        sliceData.Add("slice_id", cellValue == null ? "" : cellValue.ToString());
                                        continue;
                                    case 2:
                                        sliceData.Add("slice_number", cellValue == null ? "" : cellValue.ToString());
                                        continue;
                                    case 3:
                                        sliceData.Add("die_quantity", cellValue == null ? "" : cellValue.ToString());
                                        sliceList.Add(sliceData);
                                        continue;
                                    default:
                                        continue;
                                }
                            }
                            break;
                    }
                }
                dataList.Add(rowData);
                //sliceList.Add(sliceData);
            }
            using var connection = new SqlConnection(CommonFunction.connectionString);
            connection.Open();
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_raw_data";

                // Map các cột từ listRaw vào bảng đích
                bulkCopy.ColumnMappings.Add("key_no", "key_no");
                bulkCopy.ColumnMappings.Add("wafer_lot", "wafer_lot");
                bulkCopy.ColumnMappings.Add("mapping", "mapping");
                bulkCopy.ColumnMappings.Add("source_device", "source_device");
                bulkCopy.ColumnMappings.Add("wafer_quantity", "wafer_quantity");
                bulkCopy.ColumnMappings.Add("die_quantity", "die_quantity");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                // Tạo DataTable từ listRaw
                var dataTable = new DataTable();
                dataTable.Columns.Add("key_no", typeof(string));
                dataTable.Columns.Add("wafer_lot", typeof(string));
                dataTable.Columns.Add("mapping", typeof(string));
                dataTable.Columns.Add("source_device", typeof(string));
                dataTable.Columns.Add("wafer_quantity", typeof(string));
                dataTable.Columns.Add("die_quantity", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in dataList)
                {
                    var row = dataTable.NewRow();
                    row["key_no"] = item["key_no"];
                    row["wafer_lot"] = item["wafer_lot"];
                    row["mapping"] = item["mapping"];
                    row["source_device"] = item["source_device"];
                    row["wafer_quantity"] = item["wafer_quantity"];
                    row["die_quantity"] = item["die_quantity"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                // Sử dụng SqlBulkCopy để lưu dữ liệu vào cơ sở dữ liệu
                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }

            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_slice";

                bulkCopy.ColumnMappings.Add("slice_id", "slice_id");
                bulkCopy.ColumnMappings.Add("slice_number", "slice_number");
                bulkCopy.ColumnMappings.Add("die_quantity", "die_quantity");
                bulkCopy.ColumnMappings.Add("wafer_lot", "wafer_lot");
                bulkCopy.ColumnMappings.Add("thickness", "thickness");
                bulkCopy.ColumnMappings.Add("backgroundColor", "backgroundColor");
                bulkCopy.ColumnMappings.Add("die_alloc", "die_alloc");
                bulkCopy.ColumnMappings.Add("wafer_lot_dcc", "wafer_lot_dcc");
                bulkCopy.ColumnMappings.Add("remain", "remain");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                var dataTable = new DataTable();
                dataTable.Columns.Add("slice_id", typeof(string));
                dataTable.Columns.Add("slice_number", typeof(string));
                dataTable.Columns.Add("die_quantity", typeof(string));
                dataTable.Columns.Add("wafer_lot", typeof(string));
                dataTable.Columns.Add("thickness", typeof(string));
                dataTable.Columns.Add("backgroundColor", typeof(string));
                dataTable.Columns.Add("die_alloc", typeof(string));
                dataTable.Columns.Add("wafer_lot_dcc", typeof(string));
                dataTable.Columns.Add("remain", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in sliceList)
                {
                    var row = dataTable.NewRow();
                    row["slice_id"] = item["slice_id"];
                    row["slice_number"] = item["slice_number"];
                    row["die_quantity"] = item["die_quantity"];
                    row["wafer_lot"] = item["wafer_lot"];
                    row["thickness"] = item["thickness"];
                    row["backgroundColor"] = item["backgroundColor"];
                    row["die_alloc"] = item["die_alloc"];
                    row["wafer_lot_dcc"] = item["wafer_lot_dcc"];
                    row["remain"] = item["remain"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }
            
        }
        private void UploadWaferLot(ExcelWorksheet worksheet, string userid)
        {
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                if (row == 1 || row == 2) continue;
                Dictionary<string, string> rowData = new Dictionary<string, string>();
                rowData.Add("id_card", userid);

                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    // Lấy giá trị của cell
                    var cellValue = worksheet.Cells[row, col].Value;
                    switch (col)
                    {
                        case 4:
                            rowData.Add("source_device", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 5:
                            rowData.Add("material", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 6:
                            rowData.Add("batch", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 7:
                            rowData.Add("customer_lot", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 11:
                            rowData.Add("invoice_wafer_qty", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 28:
                            rowData.Add("wafer_mapping", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 15:
                            rowData.Add("receive_date", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 39:
                            rowData.Add("key_no", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        default:
                            continue;
                    }
                }
                dataList.Add(rowData);
            }
            using var connection = new SqlConnection(CommonFunction.connectionString);
            connection.Open();
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_wafer_lot";

                // Map các cột từ listRaw vào bảng đích
                bulkCopy.ColumnMappings.Add("source_device", "source_device");
                bulkCopy.ColumnMappings.Add("material", "material");
                bulkCopy.ColumnMappings.Add("batch", "batch");
                bulkCopy.ColumnMappings.Add("customer_lot", "customer_lot");
                bulkCopy.ColumnMappings.Add("invoice_wafer_qty", "invoice_wafer_qty");
                bulkCopy.ColumnMappings.Add("wafer_mapping", "wafer_mapping");
                bulkCopy.ColumnMappings.Add("receive_date", "receive_date");
                bulkCopy.ColumnMappings.Add("key_no", "key_no");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                // Tạo DataTable từ listRaw
                var dataTable = new DataTable();
                dataTable.Columns.Add("source_device", typeof(string));
                dataTable.Columns.Add("material", typeof(string));
                dataTable.Columns.Add("batch", typeof(string));
                dataTable.Columns.Add("customer_lot", typeof(string));
                dataTable.Columns.Add("invoice_wafer_qty", typeof(string));
                dataTable.Columns.Add("wafer_mapping", typeof(string));
                dataTable.Columns.Add("receive_date", typeof(string));
                dataTable.Columns.Add("key_no", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in dataList)
                {
                    var row = dataTable.NewRow();
                    row["source_device"] = item["source_device"];
                    row["material"] = item["material"];
                    row["batch"] = item["batch"];
                    row["customer_lot"] = item["customer_lot"];
                    row["invoice_wafer_qty"] = item["invoice_wafer_qty"];
                    row["wafer_mapping"] = item["wafer_mapping"];
                    row["receive_date"] = item["receive_date"];
                    row["key_no"] = item["key_no"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                // Sử dụng SqlBulkCopy để lưu dữ liệu vào cơ sở dữ liệu
                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }
        }
        private void UploadHeader1(ExcelWorksheet worksheet, string userid)
        {
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                if (col == 1) continue;
                Dictionary<string, string> colData = new Dictionary<string, string>();
                colData.Add("id_card", userid);

                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Lấy giá trị của cell
                    var cellValue = worksheet.Cells[row, col].Value;
                    switch (row)
                    {
                        case 1:
                            colData.Add("target_device", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 2:
                            colData.Add("pdl", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 3:
                            colData.Add("marking", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 4:
                            colData.Add("bd", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 5:
                            colData.Add("body_size", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 6:
                            colData.Add("pcb", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 7:
                            colData.Add("wafer_name", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 8:
                            colData.Add("mapping", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 9:
                            colData.Add("po", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 10:
                            colData.Add("ship_to", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 11:
                            colData.Add("npi_flag", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 12:
                            colData.Add("lot_type", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 13:
                            colData.Add("device", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 14:
                            colData.Add("cus_info", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        default:
                            continue;
                    }
                }
                dataList.Add(colData);
            }
            using var connection = new SqlConnection(CommonFunction.connectionString);
            connection.Open();
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_header1";

                // Map các cột từ listRaw vào bảng đích
                bulkCopy.ColumnMappings.Add("target_device", "target_device");
                bulkCopy.ColumnMappings.Add("pdl", "pdl");
                bulkCopy.ColumnMappings.Add("marking", "marking");
                bulkCopy.ColumnMappings.Add("bd", "bd");
                bulkCopy.ColumnMappings.Add("body_size", "body_size");
                bulkCopy.ColumnMappings.Add("pcb", "pcb");
                bulkCopy.ColumnMappings.Add("wafer_name", "wafer_name");
                bulkCopy.ColumnMappings.Add("mapping", "mapping");
                bulkCopy.ColumnMappings.Add("po", "po");
                bulkCopy.ColumnMappings.Add("ship_to", "ship_to");
                bulkCopy.ColumnMappings.Add("npi_flag", "npi_flag");
                bulkCopy.ColumnMappings.Add("lot_type", "lot_type");
                bulkCopy.ColumnMappings.Add("device", "device");
                bulkCopy.ColumnMappings.Add("cus_info", "cus_info");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                // Tạo DataTable từ listRaw
                var dataTable = new DataTable();
                dataTable.Columns.Add("target_device", typeof(string));
                dataTable.Columns.Add("pdl", typeof(string));
                dataTable.Columns.Add("marking", typeof(string));
                dataTable.Columns.Add("bd", typeof(string));
                dataTable.Columns.Add("body_size", typeof(string));
                dataTable.Columns.Add("pcb", typeof(string));
                dataTable.Columns.Add("wafer_name", typeof(string));
                dataTable.Columns.Add("mapping", typeof(string));
                dataTable.Columns.Add("po", typeof(string));
                dataTable.Columns.Add("ship_to", typeof(string));
                dataTable.Columns.Add("npi_flag", typeof(string));
                dataTable.Columns.Add("lot_type", typeof(string));
                dataTable.Columns.Add("device", typeof(string));
                dataTable.Columns.Add("cus_info", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in dataList)
                {
                    var row = dataTable.NewRow();
                    row["target_device"] = item["target_device"];
                    row["pdl"] = item["pdl"];
                    row["marking"] = item["marking"];
                    row["bd"] = item["bd"];
                    row["body_size"] = item["body_size"];
                    row["pcb"] = item["pcb"];
                    row["wafer_name"] = item["wafer_name"];
                    row["mapping"] = item["mapping"];
                    row["po"] = item["po"];
                    row["ship_to"] = item["ship_to"];
                    row["npi_flag"] = item["npi_flag"];
                    row["lot_type"] = item["lot_type"];
                    row["device"] = item["device"];
                    row["cus_info"] = item["cus_info"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                // Sử dụng SqlBulkCopy để lưu dữ liệu vào cơ sở dữ liệu
                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }
        }
        private void UploadHeader2(ExcelWorksheet worksheet, string userid)
        {
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                if (row == 1) continue;
                Dictionary<string, string> rowData = new Dictionary<string, string>();
                rowData.Add("id_card", userid);

                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    // Lấy giá trị của cell
                    var cellValue = worksheet.Cells[row, col].Value;
                    switch (col)
                    {
                        case 1:
                            rowData.Add("factory", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 2:
                            rowData.Add("fg", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 3:
                            rowData.Add("pv", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 4:
                            rowData.Add("ct", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 5:
                            rowData.Add("line", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 6:
                            rowData.Add("ship_to", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        case 7:
                            rowData.Add("status", cellValue == null ? "" : cellValue.ToString());
                            continue;
                        default:
                            continue;
                    }
                }
                dataList.Add(rowData);
            }
            using var connection = new SqlConnection(CommonFunction.connectionString);
            connection.Open();
            using (var bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "da_header2";

                // Map các cột từ listRaw vào bảng đích
                bulkCopy.ColumnMappings.Add("factory", "factory");
                bulkCopy.ColumnMappings.Add("fg", "fg");
                bulkCopy.ColumnMappings.Add("pv", "pv");
                bulkCopy.ColumnMappings.Add("ct", "ct");
                bulkCopy.ColumnMappings.Add("line", "line");
                bulkCopy.ColumnMappings.Add("ship_to", "ship_to");
                bulkCopy.ColumnMappings.Add("status", "status");
                bulkCopy.ColumnMappings.Add("id_card", "id_card");

                // Tạo DataTable từ listRaw
                var dataTable = new DataTable();
                dataTable.Columns.Add("factory", typeof(string));
                dataTable.Columns.Add("fg", typeof(string));
                dataTable.Columns.Add("pv", typeof(string));
                dataTable.Columns.Add("ct", typeof(string));
                dataTable.Columns.Add("line", typeof(string));
                dataTable.Columns.Add("ship_to", typeof(string));
                dataTable.Columns.Add("status", typeof(string));
                dataTable.Columns.Add("id_card", typeof(string));

                foreach (var item in dataList)
                {
                    var row = dataTable.NewRow();
                    row["factory"] = item["factory"];
                    row["fg"] = item["fg"];
                    row["pv"] = item["pv"];
                    row["ct"] = item["ct"];
                    row["line"] = item["line"];
                    row["ship_to"] = item["ship_to"];
                    row["status"] = item["status"];
                    row["id_card"] = item["id_card"];
                    dataTable.Rows.Add(row);
                }

                // Sử dụng SqlBulkCopy để lưu dữ liệu vào cơ sở dữ liệu
                bulkCopy.WriteToServer(dataTable);
                bulkCopy.Close();
            }
        }

    }
}
