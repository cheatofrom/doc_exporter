import { FormEvent, useMemo, useState } from "react";
import { ReportFormData } from "./types";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:39011";

const defaultFormData: ReportFormData = {
  // 第一页默认值
  report_no: "",
  project_name: "EA3800N 项目",
  report_title: "试验报告",
  subject: "",
  prepared_by: "",
  prepared_date: "",
  reviewed_by: "",
  reviewed_date: "",
  approved_by: "",
  approved_date: "",
  company_name: "杭州时代电动科技有限公司",

  // 第二页默认值
  sample_name: "EA3800N 整桥",
  sample_project: "EA3800N",
  production_stage: "B样",
  sample_quantity: "1",
  sample_no: "EA3800N101-202500417-001",
  test_location: "杭州时代电动试验室",
  send_date: "2025/4/17",
  test_date: "2025/4/28",
  task_source: "试验申请单",
  test_item: "整桥温升试验",
  test_basis_method: `测试工况1：冷却水温：25℃，油泵转速2500rpm，透明盖，流量23L/min电机旋变处增加进油管。（二挡）
7000rpm-200Nm、9550rpm-200Nm、9550rpm-0Nm
测试工况2：冷却水温：25℃，油泵转速2500rpm，原前盖，流量23L/min电机旋变处增加进油管。（二挡）
7000rpm-200Nm
测试工况3：冷却水温：65℃，油泵转速3800rpm，原前盖，流量23L/min电机旋变处增加进油管。（二挡）
9550rpm-200Nm
测试工况4：冷却水温：25℃，油泵转速2500rpm，流量23L/min（峰值温升测试）（二挡）
6221rpm-460Nm、13000rpm-220Nm`,
  test_basis_criteria: "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）",
  test_conclusion: `整桥温升试验：
测试工况1：9550rpm-0Nm、7000rpm-200Nm工况（符合） 9550rpm-200Nm工况（不符合）
测试工况2：7000rpm-200Nm工况（符合）
测试工况3：9550rpm-200Nm（不符合）
测试工况4：峰值温升不评判，结果如8试验结果`,
  test_attachments: `1. 试验目的：见附录A
2. 试验对象：见附录B
3. 试验条件：见附录C
4. 试验步骤：见附录D
5. 试验结果：见附录E
6. 拆解结果：见附录F`,

  // 第三页默认值
  test_purpose: "验证整桥在二挡额定工况和峰值工况下的电机与油温最高温度点。",
  assembly_list: [
    { id: 1, name: "EA3800N 整桥", part_no: "EA3800N101-20250417-001", version: "/", batch_no: "/", remark: "/" }
  ],
  component_list: [
    { id: 1, name: "油泵总成-左", part_no: "P400N1301110", version: "AAA", trace_no: "/", supplier: "湖南腾智机电有限责任公司", remark: "" },
    { id: 2, name: "吸滤总成", part_no: "P380N1320010", version: "AAA", trace_no: "/", supplier: "浙江环球滤清器有限公司", remark: "" },
    { id: 3, name: "吸滤底座", part_no: "P380N1320040", version: "AAA", trace_no: "202411060050100000", supplier: "杭州大杉科技有限公司", remark: "" }
  ],
  photo_list: [],

  // 附录 B4 默认值
  motor_manufacturer: "龙芯电驱动",
  motor_model: "TZ220XYC01",
  motor_type: "永磁同步电机",
  motor_rated_efficiency: "≥96%",
  motor_peak_power: "300",
  motor_rated_power: "200",
  motor_max_torque: "460",
  motor_rated_torque: "200",
  motor_max_speed: "13000",
  motor_rated_speed: "9550",
  motor_cooling_method: "油冷",
  motor_water_pipe_dia: "15",
  motor_water_inlet_temp: "25±2",
  motor_water_flow: "≥15L/min",
  motor_test_state: "批量件",
  motor_steel_stamp: "/",

  ctrl_manufacturer: "法拉第科技有限公司",
  ctrl_model: "FED200-54Z232S-HZSD6",
  ctrl_continuous_current: "246A(AC)",
  ctrl_work_system: "S9",
  ctrl_short_current: "550A",
  ctrl_phase_num: "三相",
  ctrl_cooling_method: "水冷",
  ctrl_factory_no: "59120334G1020004",
  ctrl_software_program: `FED20_TM_B1.08M184S1V0.02_250201_Encrypt3_250601_Encrypt3.mot16
FED20_ISG_B1.08M184S1V0.02_250201_Encrypt3_250601_Encrypt3.mot16`,

  bridge_manufacturer: "CDTL",
  bridge_model: "EA3800N",
  bridge_voltage: "630V",
  bridge_max_output_torque: "38511Nm",
  bridge_rated_power: "200kw",
  bridge_peak_power: "300kw",
  bridge_cooling_method: "系统内油冷 系统外水冷",
  bridge_gear_ratio: "83.72/21.4",
  bridge_max_output_speed: "607",
  bridge_rated_load: "13T",

  // 附录 C 默认值
  test_leader: "王佳忠",
  test_engineer: "王佳忠",
  c_cooling_method: "水冷+油冷",
  c_cooling_pressure_flow: "23L/min",
  c_cooling_water_temp: "65℃/25℃",
  test_time_list: [
    { id: 1, stage: "温升试验", time: "2025/4/28-2025/4/29", env_temp: "27℃", env_humidity: "56%RH", location: "杭州时代电动试验室" }
  ],
  equipment_list: [
    { id: 1, name: "电源柜", equip_no: "CDTL-EQ-0002", quantity: "1", valid_date: "无需校验", remark_image: "" }
  ],

  // 附录 D 默认值
  test_steps: `a. 测试桥固定在台架上，将两侧轮边分别用联轴器连接并装上负载。
b. 安装水管，三相线，电机旋变插件，油泵插件，气缸插件并通上气，不锁差速锁。
c. 布置温度采集仪，将定子温度线连接至采集仪。
d. 将水温调至25℃，将桥挂上二挡。
e. 一挡额定温升：两负载同时给转速，中间给扭矩，记录电机、油温数据
f. 测试工况1：冷却水温：25℃，油泵转速2500rpm，透明盖，流量23L/min电机旋变处增加进油管。7000rpm-200Nm、9550rpm-200Nm、9550rpm-0Nm。
g. 测试工况2：冷却水温：25℃，油泵转速2500rpm，原前盖，流量23L/min电机旋变处增加进油管。7000rpm-200Nm。
h. 测试工况3：冷却水温：65℃，油泵转速3800rpm，原前盖，流量23L/min电机旋变处增加进油管。9550rpm-200Nm。
i. 测试工况4：冷却水温：25℃，油泵转速2500rpm，流量23L/min（峰值温升测试）（二挡）6221rpm-460Nm、13000rpm-220Nm。`,

  // 附录 E1 默认值
  e1_test_results: [
    {
      id: 1,
      test_item: "整桥温升试验",
      test_condition: "冷却水温：25℃ 油泵转速2500rpm 透明盖 流量23L/mi 电机旋变处增加进油管。",
      evaluation_criteria: "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）",
      test_result: "7000rpm-200Nm：\n运行51min，油温66℃，电机温度130℃；已平衡\n9550rpm-200Nm：\n运行32min，油温76℃，电机＞150℃。\n9550rpm-0Nm：\n运行18min，油温49℃，电机64℃。已平衡",
      compliance: "不符合"
    },
    {
      id: 2,
      test_item: "整桥温升试验",
      test_condition: "冷却水温：25℃ 油泵转速2500rpm 原前盖 流量23L/min 电机旋变处增加进油管。",
      evaluation_criteria: "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）",
      test_result: "7000rpm-200Nm：\n运行50min，油温68℃，电机122℃。已平衡",
      compliance: "不符合"
    },
    {
      id: 3,
      test_item: "整桥温升试验",
      test_condition: "冷却水温：65℃ 油泵转速3800rpm原前盖 流量23L/min 电机旋变处增加进油管。",
      evaluation_criteria: "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）",
      test_result: "9550pm-200Nm：\n运行21min，油温103℃，电机温度＞150℃。",
      compliance: "不符合"
    },
    {
      id: 4,
      test_item: "整桥温升试验",
      test_condition: "冷却水温：25℃，油泵转速2500rpm，流量23L/min",
      evaluation_criteria: "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）",
      test_result: "6221rpm-460Nm：\n运行58s，油温44℃，电机温度＞150℃\n13000rpm-220Nm：\n运行96s，油温49℃，电机温度＞150℃",
      compliance: "/"
    }
  ],

  // 附录 E2 默认值
  e2_charts: [],

  // 附录 E3 默认值
  e3_issues: [
    {
      id: 1,
      occur_time: "2025/4/29",
      description: "测试工况1：\n9550rpm-200Nm：运行32min，油温76℃，电机＞150℃。",
      responsible_party: "设计",
      solution: "无",
      is_resolved: "否"
    },
    {
      id: 2,
      occur_time: "2025/4/29",
      description: "测试工况3：\n9550pm-200Nm：运行21min，油温103℃，电机温度＞150℃。",
      responsible_party: "设计",
      solution: "无",
      is_resolved: "否"
    },
    {
      id: 3,
      occur_time: "/",
      description: "无",
      responsible_party: "无",
      solution: "无",
      is_resolved: "否"
    }
  ],

  // 附录 E4 默认值
  e4_issue_photos: [],

  // 附录 F1 默认值
  f1_teardown_results: [
    { id: 1, part_name: "/", part_no: "/", evaluation_criteria: "/", teardown_result: "/", compliance: "/", remark: "/" },
    { id: 2, part_name: "/", part_no: "/", evaluation_criteria: "/", teardown_result: "/", compliance: "/", remark: "/" },
    { id: 3, part_name: "/", part_no: "/", evaluation_criteria: "/", teardown_result: "/", compliance: "/", remark: "/" },
    { id: 4, part_name: "/", part_no: "/", evaluation_criteria: "/", teardown_result: "/", compliance: "/", remark: "/" },
    { id: 5, part_name: "/", part_no: "/", evaluation_criteria: "/", teardown_result: "/", compliance: "/", remark: "/" }
  ],

  // 附录 F2 默认值
  f2_teardown_issues: [
    { id: 1, occur_time: "/", issue_description: "/", description: "/", remark: "/" },
    { id: 2, occur_time: "/", issue_description: "/", description: "/", remark: "/" },
    { id: 3, occur_time: "/", issue_description: "/", description: "/", remark: "/" }
  ],

  // 附录 F3 默认值
  f3_teardown_photos: []
};

function App() {
  const [formData, setFormData] = useState<ReportFormData>(defaultFormData);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [uploadMessage, setUploadMessage] = useState("");
  const [isFirstPageOpen, setIsFirstPageOpen] = useState(false);
  const [isSecondPageOpen, setIsSecondPageOpen] = useState(false);
  const [isThirdPageOpen, setIsThirdPageOpen] = useState(false);
  const [isFourthPageOpen, setIsFourthPageOpen] = useState(false);
  const [isFifthPageOpen, setIsFifthPageOpen] = useState(false);
  const [isSixthPageOpen, setIsSixthPageOpen] = useState(false);
  const [isSeventhPageOpen, setIsSeventhPageOpen] = useState(false);
  const [isEighthPageOpen, setIsEighthPageOpen] = useState(false);
  const [isNinthPageOpen, setIsNinthPageOpen] = useState(false);
  const [isTenthPageOpen, setIsTenthPageOpen] = useState(true);

  const requiredFields = useMemo(
    () => [
      formData.report_no,
      formData.project_name,
      formData.report_title,
      formData.subject,
      formData.prepared_by,
      formData.prepared_date,
      formData.reviewed_by,
      formData.reviewed_date,
      formData.approved_by,
      formData.approved_date,
      formData.company_name,
      // 第二页字段必填校验
      formData.sample_name,
      formData.sample_project,
      formData.production_stage,
      formData.sample_quantity,
      formData.sample_no,
      formData.test_location,
      formData.send_date,
      formData.test_date,
      formData.task_source,
      formData.test_item,
      formData.test_basis_method,
      formData.test_basis_criteria,
      formData.test_conclusion,
      formData.test_attachments,
      // 第三页字段
      formData.test_purpose
    ],
    [formData]
  );

  const isValid = requiredFields.every((fieldValue) => (fieldValue || "").trim().length > 0);

  const handleChange = (key: keyof ReportFormData, value: string | any[]) => {
    setFormData((prev) => ({
      ...prev,
      [key]: value
    }));
  };

  const handleAssemblyChange = (index: number, key: string, value: string) => {
    const newList = [...formData.assembly_list];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("assembly_list", newList);
  };

  const addAssemblyRow = () => {
    const newList = [...formData.assembly_list];
    newList.push({
      id: newList.length + 1,
      name: "",
      part_no: "",
      version: "",
      batch_no: "",
      remark: ""
    });
    handleChange("assembly_list", newList);
  };

  const removeAssemblyRow = (index: number) => {
    const newList = formData.assembly_list.filter((_, i) => i !== index);
    // 重新排序序号
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("assembly_list", reorderedList);
  };

  const handleComponentChange = (index: number, key: string, value: string) => {
    const newList = [...formData.component_list];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("component_list", newList);
  };

  const addComponentRow = () => {
    const newList = [...formData.component_list];
    newList.push({
      id: newList.length + 1,
      name: "",
      part_no: "",
      version: "",
      trace_no: "",
      supplier: "",
      remark: ""
    });
    handleChange("component_list", newList);
  };

  const removeComponentRow = (index: number) => {
    const newList = formData.component_list.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("component_list", reorderedList);
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target && typeof event.target.result === "string") {
          setFormData((prev) => {
            const newId = prev.photo_list.length > 0 ? prev.photo_list[prev.photo_list.length - 1].id + 1 : 1;
            return {
              ...prev,
              photo_list: [
                ...prev.photo_list,
                { id: newId, label: `样件状态 ${newId}`, dataUrl: event.target!.result as string }
              ]
            };
          });
        }
      };
      reader.readAsDataURL(file);
    });
    // 清空 input 允许重复选择相同文件
    e.target.value = "";
  };

  const handlePhotoLabelChange = (id: number, newLabel: string) => {
    setFormData((prev) => ({
      ...prev,
      photo_list: prev.photo_list.map((photo) =>
        photo.id === id ? { ...photo, label: newLabel } : photo
      )
    }));
  };

  const removePhoto = (id: number) => {
    setFormData((prev) => ({
      ...prev,
      photo_list: prev.photo_list.filter((photo) => photo.id !== id)
    }));
  };

  const handleTestTimeChange = (index: number, key: string, value: string) => {
    const newList = [...formData.test_time_list];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("test_time_list", newList);
  };

  const addTestTimeRow = () => {
    const newList = [...formData.test_time_list];
    newList.push({
      id: newList.length + 1,
      stage: "",
      time: "",
      env_temp: "",
      env_humidity: "",
      location: ""
    });
    handleChange("test_time_list", newList);
  };

  const removeTestTimeRow = (index: number) => {
    const newList = formData.test_time_list.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("test_time_list", reorderedList);
  };

  const handleEquipmentChange = (index: number, key: string, value: string) => {
    const newList = [...formData.equipment_list];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("equipment_list", newList);
  };

  const handleEquipmentImageUpload = (index: number, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      if (event.target && typeof event.target.result === "string") {
        const newList = [...formData.equipment_list];
        newList[index] = { ...newList[index], remark_image: event.target.result };
        handleChange("equipment_list", newList);
      }
    };
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const addEquipmentRow = () => {
    const newList = [...formData.equipment_list];
    newList.push({
      id: newList.length + 1,
      name: "",
      equip_no: "",
      quantity: "",
      valid_date: "",
      remark_image: ""
    });
    handleChange("equipment_list", newList);
  };

  const removeEquipmentRow = (index: number) => {
    const newList = formData.equipment_list.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("equipment_list", reorderedList);
  };

  const handleE1Change = (index: number, key: string, value: string) => {
    const newList = [...formData.e1_test_results];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("e1_test_results", newList);
  };

  const addE1Row = () => {
    const newList = [...formData.e1_test_results];
    newList.push({
      id: newList.length + 1,
      test_item: newList.length > 0 ? newList[0].test_item : "整桥温升试验",
      test_condition: "",
      evaluation_criteria: newList.length > 0 ? newList[0].evaluation_criteria : "",
      test_result: "",
      compliance: ""
    });
    handleChange("e1_test_results", newList);
  };

  const removeE1Row = (index: number) => {
    const newList = formData.e1_test_results.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("e1_test_results", reorderedList);
  };

  const handleE2ChartTitleChange = (index: number, newTitle: string) => {
    const newList = [...formData.e2_charts];
    newList[index] = { ...newList[index], title: newTitle };
    handleChange("e2_charts", newList);
  };

  const removeE2Chart = (index: number) => {
    const newList = formData.e2_charts.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("e2_charts", reorderedList);
  };

  const handleE2ChartUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target && typeof event.target.result === "string") {
          setFormData((prev) => {
            const newList = [...prev.e2_charts];
            newList.push({
              id: newList.length + 1,
              title: `图 ${newList.length + 1} 试验图表`,
              chart_image: event.target?.result as string
            });
            return { ...prev, e2_charts: newList };
          });
        }
      };
      reader.readAsDataURL(file);
    });

    e.target.value = "";
  };

  const handleE3Change = (index: number, key: string, value: string) => {
    const newList = [...formData.e3_issues];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("e3_issues", newList);
  };

  const addE3Row = () => {
    const newList = [...formData.e3_issues];
    newList.push({
      id: newList.length + 1,
      occur_time: "",
      description: "",
      responsible_party: "",
      solution: "",
      is_resolved: "否"
    });
    handleChange("e3_issues", newList);
  };

  const removeE3Row = (index: number) => {
    const newList = formData.e3_issues.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("e3_issues", reorderedList);
  };

  const handleE4PhotoLabelChange = (index: number, newLabel: string) => {
    const newList = [...formData.e4_issue_photos];
    newList[index] = { ...newList[index], label: newLabel };
    handleChange("e4_issue_photos", newList);
  };

  const removeE4Photo = (index: number) => {
    const newList = formData.e4_issue_photos.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("e4_issue_photos", reorderedList);
  };

  const handleE4PhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target && typeof event.target.result === "string") {
          setFormData((prev) => {
            const newList = [...prev.e4_issue_photos];
            newList.push({
              id: newList.length + 1,
              label: `问题照片 ${newList.length + 1}`,
              dataUrl: event.target?.result as string
            });
            return { ...prev, e4_issue_photos: newList };
          });
        }
      };
      reader.readAsDataURL(file);
    });

    e.target.value = "";
  };

  const handleF1Change = (index: number, key: string, value: string) => {
    const newList = [...formData.f1_teardown_results];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("f1_teardown_results", newList);
  };

  const addF1Row = () => {
    const newList = [...formData.f1_teardown_results];
    newList.push({
      id: newList.length + 1,
      part_name: "/",
      part_no: "/",
      evaluation_criteria: "/",
      teardown_result: "/",
      compliance: "/",
      remark: "/"
    });
    handleChange("f1_teardown_results", newList);
  };

  const removeF1Row = (index: number) => {
    const newList = formData.f1_teardown_results.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("f1_teardown_results", reorderedList);
  };

  const handleF2Change = (index: number, key: string, value: string) => {
    const newList = [...formData.f2_teardown_issues];
    newList[index] = { ...newList[index], [key]: value };
    handleChange("f2_teardown_issues", newList);
  };

  const addF2Row = () => {
    const newList = [...formData.f2_teardown_issues];
    newList.push({
      id: newList.length + 1,
      occur_time: "/",
      issue_description: "/",
      description: "/",
      remark: "/"
    });
    handleChange("f2_teardown_issues", newList);
  };

  const removeF2Row = (index: number) => {
    const newList = formData.f2_teardown_issues.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("f2_teardown_issues", reorderedList);
  };

  const handleF3PhotoLabelChange = (index: number, newLabel: string) => {
    const newList = [...formData.f3_teardown_photos];
    newList[index] = { ...newList[index], label: newLabel };
    handleChange("f3_teardown_photos", newList);
  };

  const removeF3Photo = (index: number) => {
    const newList = formData.f3_teardown_photos.filter((_, i) => i !== index);
    const reorderedList = newList.map((item, i) => ({ ...item, id: i + 1 }));
    handleChange("f3_teardown_photos", reorderedList);
  };

  const handleF3PhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target && typeof event.target.result === "string") {
          setFormData((prev) => {
            const newList = [...prev.f3_teardown_photos];
            newList.push({
              id: newList.length + 1,
              label: `拆解照片 ${newList.length + 1}`,
              dataUrl: event.target?.result as string
            });
            return { ...prev, f3_teardown_photos: newList };
          });
        }
      };
      reader.readAsDataURL(file);
    });

    e.target.value = "";
  };

  const exportWord = async (event: FormEvent) => {
    event.preventDefault();
    setErrorMessage("");
    if (!isValid) {
      setErrorMessage("请填写完整后再导出");
      return;
    }

    setIsSubmitting(true);
    try {
      let response;

      if (templateFile) {
        const formDataToSend = new FormData();
        formDataToSend.append("template", templateFile);
        formDataToSend.append("data", JSON.stringify(formData));

        response = await fetch(`${API_BASE_URL}/api/export-word-custom`, {
          method: "POST",
          body: formDataToSend
        });
      } else {
        response = await fetch(`${API_BASE_URL}/api/export-word`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(formData)
        });
      }

      if (!response.ok) {
        const errorResult = await response.json().catch(() => ({ message: "导出失败" }));
        throw new Error(errorResult.message || "导出失败");
      }

      const blob = await response.blob();
      const fileName = `${formData.report_no || "试验报告"}.docx`;
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = fileName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "导出失败");
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <main className="page">
      <h1>试验报告在线填单</h1>
      <p className="subtitle">封面页字段填写后可直接导出 Word</p>

      <div className="upload-section">
        <h3>自定义模板（可选）</h3>
        <p className="upload-hint">如果不上传，将使用系统默认模板。模板内请使用 <code>[[字段名]]</code> 作为占位符。</p>
        <input 
          type="file" 
          accept=".docx" 
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) {
              setTemplateFile(file);
              setUploadMessage(`已选择文件: ${file.name}`);
            } else {
              setTemplateFile(null);
              setUploadMessage("");
            }
          }} 
        />
        {uploadMessage && <p className="upload-message">{uploadMessage}</p>}
      </div>

      <form className="form" onSubmit={exportWord}>
        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsFirstPageOpen(!isFirstPageOpen)}
          >
            <h2>第一页（封面）</h2>
            <span className={`arrow ${isFirstPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isFirstPageOpen && (
            <div className="section-content">
              <label>
                编号 <span className="field-hint">(report_no)</span>
                <input
                  value={formData.report_no}
                  onChange={(event) => handleChange("report_no", event.target.value)}
                  placeholder="例如 EA300N-SYG-2025-001"
                  required
                />
              </label>

              <label>
                项目名称 <span className="field-hint">(project_name)</span>
                <input
                  value={formData.project_name}
                  onChange={(event) => handleChange("project_name", event.target.value)}
                  required
                />
              </label>

              <label>
                文档主标题 <span className="field-hint">(report_title)</span>
                <input
                  value={formData.report_title}
                  onChange={(event) => handleChange("report_title", event.target.value)}
                  required
                />
              </label>

              <label>
                题目 <span className="field-hint">(subject)</span>
                <input
                  value={formData.subject}
                  onChange={(event) => handleChange("subject", event.target.value)}
                  placeholder="例如 EA3800N 整系温升试验报告（第三轮）"
                  required
                />
              </label>

              <label>
                编制签字 <span className="field-hint">(prepared_by)</span>
                <input
                  value={formData.prepared_by}
                  onChange={(event) => handleChange("prepared_by", event.target.value)}
                  required
                />
              </label>
              <label>
                编制日期 <span className="field-hint">(prepared_date)</span>
                <input
                  type="date"
                  value={formData.prepared_date}
                  onChange={(event) => handleChange("prepared_date", event.target.value)}
                  required
                />
              </label>

              <label>
                审核签字 <span className="field-hint">(reviewed_by)</span>
                <input
                  value={formData.reviewed_by}
                  onChange={(event) => handleChange("reviewed_by", event.target.value)}
                  required
                />
              </label>
              <label>
                审核日期 <span className="field-hint">(reviewed_date)</span>
                <input
                  type="date"
                  value={formData.reviewed_date}
                  onChange={(event) => handleChange("reviewed_date", event.target.value)}
                  required
                />
              </label>

              <label>
                批准签字 <span className="field-hint">(approved_by)</span>
                <input
                  value={formData.approved_by}
                  onChange={(event) => handleChange("approved_by", event.target.value)}
                  required
                />
              </label>
              <label>
                批准日期 <span className="field-hint">(approved_date)</span>
                <input
                  type="date"
                  value={formData.approved_date}
                  onChange={(event) => handleChange("approved_date", event.target.value)}
                  required
                />
              </label>

              <label>
                公司名称 <span className="field-hint">(company_name)</span>
                <input
                  value={formData.company_name}
                  onChange={(event) => handleChange("company_name", event.target.value)}
                  required
                />
              </label>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsSecondPageOpen(!isSecondPageOpen)}
          >
            <h2>第二页（基本信息与试验总结）</h2>
            <span className={`arrow ${isSecondPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isSecondPageOpen && (
            <div className="section-content">
              <label>
                样件名称 <span className="field-hint">(sample_name)</span>
                <input
                  value={formData.sample_name || ""}
                  onChange={(event) => handleChange("sample_name", event.target.value)}
                  required
                />
              </label>

              <label>
                样品项目 <span className="field-hint">(sample_project)</span>
                <input
                  value={formData.sample_project || ""}
                  onChange={(event) => handleChange("sample_project", event.target.value)}
                  required
                />
              </label>

              <label>
                生产阶段 <span className="field-hint">(production_stage)</span>
                <input
                  value={formData.production_stage || ""}
                  onChange={(event) => handleChange("production_stage", event.target.value)}
                  required
                />
              </label>

              <label>
                样品数量 <span className="field-hint">(sample_quantity)</span>
                <input
                  value={formData.sample_quantity || ""}
                  onChange={(event) => handleChange("sample_quantity", event.target.value)}
                  required
                />
              </label>

              <label>
                样件编号 <span className="field-hint">(sample_no)</span>
                <input
                  value={formData.sample_no || ""}
                  onChange={(event) => handleChange("sample_no", event.target.value)}
                  required
                />
              </label>

              <label>
                检测地点 <span className="field-hint">(test_location)</span>
                <input
                  value={formData.test_location || ""}
                  onChange={(event) => handleChange("test_location", event.target.value)}
                  required
                />
              </label>

              <label>
                送样时间 <span className="field-hint">(send_date)</span>
                <input
                  value={formData.send_date || ""}
                  onChange={(event) => handleChange("send_date", event.target.value)}
                  required
                />
              </label>

              <label>
                试验时间 <span className="field-hint">(test_date)</span>
                <input
                  value={formData.test_date || ""}
                  onChange={(event) => handleChange("test_date", event.target.value)}
                  required
                />
              </label>

              <label>
                任务来源 <span className="field-hint">(task_source)</span>
                <input
                  value={formData.task_source || ""}
                  onChange={(event) => handleChange("task_source", event.target.value)}
                  required
                />
              </label>

              <label>
                试验项目 <span className="field-hint">(test_item)</span>
                <input
                  value={formData.test_item || ""}
                  onChange={(event) => handleChange("test_item", event.target.value)}
                  required
                />
              </label>

              <label className="full-width">
                试验方法（多行） <span className="field-hint">(test_method)</span>
                <textarea
                  value={formData.test_basis_method || ""}
                  onChange={(event) => handleChange("test_basis_method", event.target.value)}
                  rows={8}
                  required
                />
              </label>

              <label className="full-width">
                评价标准（多行） <span className="field-hint">(test_criteria)</span>
                <textarea
                  value={formData.test_basis_criteria || ""}
                  onChange={(event) => handleChange("test_basis_criteria", event.target.value)}
                  rows={3}
                  required
                />
              </label>

              <label className="full-width">
                试验结论（多行） <span className="field-hint">(test_conclusion)</span>
                <textarea
                  value={formData.test_conclusion || ""}
                  onChange={(event) => handleChange("test_conclusion", event.target.value)}
                  rows={6}
                  required
                />
              </label>

              <label className="full-width">
                附录说明（多行） <span className="field-hint">(appendix_notes)</span>
                <textarea
                  value={formData.test_attachments || ""}
                  onChange={(event) => handleChange("test_attachments", event.target.value)}
                  rows={6}
                  required
                />
              </label>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsThirdPageOpen(!isThirdPageOpen)}
          >
            <h2>附录（试验目的与对象）</h2>
            <span className={`arrow ${isThirdPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isThirdPageOpen && (
            <div className="section-content">
              <label className="full-width">
                试验目的（附录A） <span className="field-hint">(test_purpose)</span>
                <textarea
                  value={formData.test_purpose || ""}
                  onChange={(event) => handleChange("test_purpose", event.target.value)}
                  rows={2}
                  required
                />
              </label>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>B1 样件总成信息</h3>
                  <button type="button" className="add-btn" onClick={addAssemblyRow}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '60px' }}>序号</th>
                        <th>测试样件名称 <span className="field-hint">(name)</span></th>
                        <th>试验样件号 <span className="field-hint">(part_no)</span></th>
                        <th>版本号 <span className="field-hint">(version)</span></th>
                        <th>批次号 <span className="field-hint">(batch_no)</span></th>
                        <th>备注 <span className="field-hint">(remark)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.assembly_list.map((row, index) => (
                        <tr key={`assembly-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.name} onChange={e => handleAssemblyChange(index, "name", e.target.value)} /></td>
                          <td><input value={row.part_no} onChange={e => handleAssemblyChange(index, "part_no", e.target.value)} /></td>
                          <td><input value={row.version} onChange={e => handleAssemblyChange(index, "version", e.target.value)} /></td>
                          <td><input value={row.batch_no} onChange={e => handleAssemblyChange(index, "batch_no", e.target.value)} /></td>
                          <td><input value={row.remark} onChange={e => handleAssemblyChange(index, "remark", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeAssemblyRow(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>B2 样件零部件信息</h3>
                  <button type="button" className="add-btn" onClick={addComponentRow}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '60px' }}>序号</th>
                        <th>测试零部件名称 <span className="field-hint">(name/part_name)</span></th>
                        <th>零件号 <span className="field-hint">(part_no)</span></th>
                        <th>版本号 <span className="field-hint">(version)</span></th>
                        <th>追溯号 <span className="field-hint">(trace_no)</span></th>
                        <th>供应商 <span className="field-hint">(supplier)</span></th>
                        <th>备注 <span className="field-hint">(remark)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.component_list.map((row, index) => (
                        <tr key={`component-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.name} onChange={e => handleComponentChange(index, "name", e.target.value)} /></td>
                          <td><input value={row.part_no} onChange={e => handleComponentChange(index, "part_no", e.target.value)} /></td>
                          <td><input value={row.version} onChange={e => handleComponentChange(index, "version", e.target.value)} /></td>
                          <td><input value={row.trace_no} onChange={e => handleComponentChange(index, "trace_no", e.target.value)} /></td>
                          <td><input value={row.supplier} onChange={e => handleComponentChange(index, "supplier", e.target.value)} /></td>
                          <td><input value={row.remark} onChange={e => handleComponentChange(index, "remark", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeComponentRow(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>B3 样件照片</h3>
                  <div className="upload-btn-wrapper">
                    <button type="button" className="add-btn">+ 添加照片</button>
                    <input 
                      type="file" 
                      accept="image/*" 
                      multiple 
                      onChange={handlePhotoUpload} 
                      className="hidden-file-input"
                    />
                  </div>
                </div>
                
                <div className="photo-grid">
                  {formData.photo_list.map((photo) => (
                    <div key={`photo-${photo.id}`} className="photo-card">
                      <button 
                        type="button" 
                        className="del-photo-btn" 
                        onClick={() => removePhoto(photo.id)}
                        title="删除照片"
                      >
                        ×
                      </button>
                      <div className="photo-img-wrapper">
                        <img src={photo.dataUrl} alt={photo.label} />
                      </div>
                      <input 
                        className="photo-label-input"
                        value={photo.label} 
                        onChange={(e) => handlePhotoLabelChange(photo.id, e.target.value)} 
                        placeholder="请输入图注..."
                      />
                    </div>
                  ))}
                  {formData.photo_list.length === 0 && (
                    <p className="upload-hint">暂无照片，请点击右上角按钮添加</p>
                  )}
                </div>
              </div>

            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsFourthPageOpen(!isFourthPageOpen)}
          >
            <h2>附录 B4 样机技术参数</h2>
            <span className={`arrow ${isFourthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isFourthPageOpen && (
            <div className="section-content">
              <div className="full-width table-section">
                <div className="table-header">
                  <h3>表 B4-1 电机参数表</h3>
                </div>
                <div className="table-container">
                  <table>
                    <tbody>
                      <tr>
                        <th>生产厂家 <span className="field-hint">(motor_manufacturer)</span></th>
                        <td><input value={formData.motor_manufacturer || ""} onChange={e => handleChange("motor_manufacturer", e.target.value)} /></td>
                        <th>型号 <span className="field-hint">(motor_model)</span></th>
                        <td><input value={formData.motor_model || ""} onChange={e => handleChange("motor_model", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>电机类型 <span className="field-hint">(motor_type)</span></th>
                        <td><input value={formData.motor_type || ""} onChange={e => handleChange("motor_type", e.target.value)} /></td>
                        <th>额定效率(%) <span className="field-hint">(motor_rated_efficiency)</span></th>
                        <td><input value={formData.motor_rated_efficiency || ""} onChange={e => handleChange("motor_rated_efficiency", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>峰值功率(kW) <span className="field-hint">(motor_peak_power)</span></th>
                        <td><input value={formData.motor_peak_power || ""} onChange={e => handleChange("motor_peak_power", e.target.value)} /></td>
                        <th>额定功率(kW) <span className="field-hint">(motor_rated_power)</span></th>
                        <td><input value={formData.motor_rated_power || ""} onChange={e => handleChange("motor_rated_power", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>最大扭矩(N·m) <span className="field-hint">(motor_max_torque)</span></th>
                        <td><input value={formData.motor_max_torque || ""} onChange={e => handleChange("motor_max_torque", e.target.value)} /></td>
                        <th>额定扭矩(N·m) <span className="field-hint">(motor_rated_torque)</span></th>
                        <td><input value={formData.motor_rated_torque || ""} onChange={e => handleChange("motor_rated_torque", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>最高转速(rpm) <span className="field-hint">(motor_max_speed)</span></th>
                        <td><input value={formData.motor_max_speed || ""} onChange={e => handleChange("motor_max_speed", e.target.value)} /></td>
                        <th>额定转速(rpm) <span className="field-hint">(motor_rated_speed)</span></th>
                        <td><input value={formData.motor_rated_speed || ""} onChange={e => handleChange("motor_rated_speed", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>冷却方式 <span className="field-hint">(motor_cooling_method)</span></th>
                        <td><input value={formData.motor_cooling_method || ""} onChange={e => handleChange("motor_cooling_method", e.target.value)} /></td>
                        <th>热交换器水管接头直径(mm) <span className="field-hint">(motor_water_pipe_dia)</span></th>
                        <td><input value={formData.motor_water_pipe_dia || ""} onChange={e => handleChange("motor_water_pipe_dia", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>冷却水入口温度(℃) <span className="field-hint">(motor_water_inlet_temp)</span></th>
                        <td><input value={formData.motor_water_inlet_temp || ""} onChange={e => handleChange("motor_water_inlet_temp", e.target.value)} /></td>
                        <th>冷却水流量(L/min) <span className="field-hint">(motor_water_flow)</span></th>
                        <td><input value={formData.motor_water_flow || ""} onChange={e => handleChange("motor_water_flow", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>被试件状态 <span className="field-hint">(motor_test_state)</span></th>
                        <td><input value={formData.motor_test_state || ""} onChange={e => handleChange("motor_test_state", e.target.value)} /></td>
                        <th>被试件钢印号 <span className="field-hint">(motor_steel_stamp)</span></th>
                        <td><input value={formData.motor_steel_stamp || ""} onChange={e => handleChange("motor_steel_stamp", e.target.value)} /></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>表 B4-2 控制器参数表</h3>
                </div>
                <div className="table-container">
                  <table>
                    <tbody>
                      <tr>
                        <th>生产厂家 <span className="field-hint">(ctrl_manufacturer)</span></th>
                        <td><input value={formData.ctrl_manufacturer || ""} onChange={e => handleChange("ctrl_manufacturer", e.target.value)} /></td>
                        <th>型号 <span className="field-hint">(ctrl_model)</span></th>
                        <td><input value={formData.ctrl_model || ""} onChange={e => handleChange("ctrl_model", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>持续工作电流(Aac) <span className="field-hint">(ctrl_continuous_current)</span></th>
                        <td><input value={formData.ctrl_continuous_current || ""} onChange={e => handleChange("ctrl_continuous_current", e.target.value)} /></td>
                        <th>工作制 <span className="field-hint">(ctrl_work_system)</span></th>
                        <td><input value={formData.ctrl_work_system || ""} onChange={e => handleChange("ctrl_work_system", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>短时工作电流 <span className="field-hint">(ctrl_short_current)</span></th>
                        <td><input value={formData.ctrl_short_current || ""} onChange={e => handleChange("ctrl_short_current", e.target.value)} /></td>
                        <th>相数 <span className="field-hint">(ctrl_phase_num)</span></th>
                        <td><input value={formData.ctrl_phase_num || ""} onChange={e => handleChange("ctrl_phase_num", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>冷却方式 <span className="field-hint">(ctrl_cooling_method)</span></th>
                        <td><input value={formData.ctrl_cooling_method || ""} onChange={e => handleChange("ctrl_cooling_method", e.target.value)} /></td>
                        <th>出厂编号 <span className="field-hint">(ctrl_factory_no)</span></th>
                        <td><input value={formData.ctrl_factory_no || ""} onChange={e => handleChange("ctrl_factory_no", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>被测控制器程序 <span className="field-hint">(ctrl_software_program)</span></th>
                        <td colSpan={3}>
                          <textarea 
                            value={formData.ctrl_software_program || ""} 
                            onChange={e => handleChange("ctrl_software_program", e.target.value)} 
                            rows={3}
                          />
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>表 B4-3 整桥参数表</h3>
                </div>
                <div className="table-container">
                  <table>
                    <tbody>
                      <tr>
                        <th>生产厂家 <span className="field-hint">(bridge_manufacturer)</span></th>
                        <td><input value={formData.bridge_manufacturer || ""} onChange={e => handleChange("bridge_manufacturer", e.target.value)} /></td>
                        <th>型号 <span className="field-hint">(bridge_model)</span></th>
                        <td><input value={formData.bridge_model || ""} onChange={e => handleChange("bridge_model", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>电压平台 <span className="field-hint">(bridge_voltage)</span></th>
                        <td><input value={formData.bridge_voltage || ""} onChange={e => handleChange("bridge_voltage", e.target.value)} /></td>
                        <th>最大输出扭矩 <span className="field-hint">(bridge_max_output_torque)</span></th>
                        <td><input value={formData.bridge_max_output_torque || ""} onChange={e => handleChange("bridge_max_output_torque", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>额定功率 <span className="field-hint">(bridge_rated_power)</span></th>
                        <td><input value={formData.bridge_rated_power || ""} onChange={e => handleChange("bridge_rated_power", e.target.value)} /></td>
                        <th>峰值功率 <span className="field-hint">(bridge_peak_power)</span></th>
                        <td><input value={formData.bridge_peak_power || ""} onChange={e => handleChange("bridge_peak_power", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>冷却方式 <span className="field-hint">(bridge_cooling_method)</span></th>
                        <td>
                          <textarea 
                            value={formData.bridge_cooling_method || ""} 
                            onChange={e => handleChange("bridge_cooling_method", e.target.value)}
                            rows={2}
                          />
                        </td>
                        <th>1档/2档速比 <span className="field-hint">(bridge_gear_ratio)</span></th>
                        <td><input value={formData.bridge_gear_ratio || ""} onChange={e => handleChange("bridge_gear_ratio", e.target.value)} /></td>
                      </tr>
                      <tr>
                        <th>最高输出转速 <span className="field-hint">(bridge_max_output_speed)</span></th>
                        <td><input value={formData.bridge_max_output_speed || ""} onChange={e => handleChange("bridge_max_output_speed", e.target.value)} /></td>
                        <th>额定轴荷 <span className="field-hint">(bridge_rated_load)</span></th>
                        <td><input value={formData.bridge_rated_load || ""} onChange={e => handleChange("bridge_rated_load", e.target.value)} /></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsFifthPageOpen(!isFifthPageOpen)}
          >
            <h2>附录 C 试验条件</h2>
            <span className={`arrow ${isFifthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isFifthPageOpen && (
            <div className="section-content">
              <label>
                试验负责人 <span className="field-hint">(test_leader)</span>
                <input
                  value={formData.test_leader || ""}
                  onChange={(e) => handleChange("test_leader", e.target.value)}
                  required
                />
              </label>

              <label>
                试验工程师 <span className="field-hint">(test_engineer)</span>
                <input
                  value={formData.test_engineer || ""}
                  onChange={(e) => handleChange("test_engineer", e.target.value)}
                  required
                />
              </label>

              <label>
                冷却方式 <span className="field-hint">(c_cooling_method)</span>
                <input
                  value={formData.c_cooling_method || ""}
                  onChange={(e) => handleChange("c_cooling_method", e.target.value)}
                  required
                />
              </label>

              <label>
                冷却水压及流量 <span className="field-hint">(c_cooling_pressure_flow)</span>
                <input
                  value={formData.c_cooling_pressure_flow || ""}
                  onChange={(e) => handleChange("c_cooling_pressure_flow", e.target.value)}
                  required
                />
              </label>

              <label className="full-width">
                冷却水温 <span className="field-hint">(c_cooling_water_temp)</span>
                <input
                  value={formData.c_cooling_water_temp || ""}
                  onChange={(e) => handleChange("c_cooling_water_temp", e.target.value)}
                  required
                />
              </label>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>C2 试验时间</h3>
                  <button type="button" className="add-btn" onClick={addTestTimeRow}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th>试验阶段 <span className="field-hint">(stage)</span></th>
                        <th>时间 <span className="field-hint">(time)</span></th>
                        <th>试验环境温度 <span className="field-hint">(env_temp)</span></th>
                        <th>试验环境湿度 <span className="field-hint">(env_humidity)</span></th>
                        <th>试验地点 <span className="field-hint">(location)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.test_time_list.map((row, index) => (
                        <tr key={`testtime-${index}`}>
                          <td><input value={row.stage} onChange={e => handleTestTimeChange(index, "stage", e.target.value)} /></td>
                          <td><input value={row.time} onChange={e => handleTestTimeChange(index, "time", e.target.value)} /></td>
                          <td><input value={row.env_temp} onChange={e => handleTestTimeChange(index, "env_temp", e.target.value)} /></td>
                          <td><input value={row.env_humidity} onChange={e => handleTestTimeChange(index, "env_humidity", e.target.value)} /></td>
                          <td><input value={row.location} onChange={e => handleTestTimeChange(index, "location", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeTestTimeRow(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section">
                <div className="table-header">
                  <h3>C3 主要测试设备</h3>
                  <button type="button" className="add-btn" onClick={addEquipmentRow}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '60px' }}>序号</th>
                        <th>设备名称 <span className="field-hint">(name)</span></th>
                        <th>设备编号 <span className="field-hint">(equip_no)</span></th>
                        <th>数量 <span className="field-hint">(quantity)</span></th>
                        <th>设备校验有效日期 <span className="field-hint">(valid_date)</span></th>
                        <th>备注(照片) <span className="field-hint">(remark_image)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.equipment_list.map((row, index) => (
                        <tr key={`equip-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.name} onChange={e => handleEquipmentChange(index, "name", e.target.value)} /></td>
                          <td><input value={row.equip_no} onChange={e => handleEquipmentChange(index, "equip_no", e.target.value)} /></td>
                          <td><input value={row.quantity} onChange={e => handleEquipmentChange(index, "quantity", e.target.value)} /></td>
                          <td><input value={row.valid_date} onChange={e => handleEquipmentChange(index, "valid_date", e.target.value)} /></td>
                          <td className="text-center">
                            {row.remark_image ? (
                              <div style={{ position: 'relative', display: 'inline-block' }}>
                                <img src={row.remark_image} alt="备注" style={{ height: '40px', objectFit: 'contain' }} />
                                <button 
                                  type="button" 
                                  style={{ position: 'absolute', top: '-5px', right: '-5px', background: 'red', color: 'white', border: 'none', borderRadius: '50%', width: '16px', height: '16px', fontSize: '10px', cursor: 'pointer', lineHeight: '16px', padding: 0 }}
                                  onClick={() => handleEquipmentChange(index, "remark_image", "")}
                                >×</button>
                              </div>
                            ) : (
                              <div className="upload-btn-wrapper">
                                <button type="button" style={{ fontSize: '12px', padding: '4px 8px' }}>上传</button>
                                <input type="file" accept="image/*" onChange={(e) => handleEquipmentImageUpload(index, e)} className="hidden-file-input" />
                              </div>
                            )}
                          </td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeEquipmentRow(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsSixthPageOpen(!isSixthPageOpen)}
          >
            <h2>附录 D 试验步骤</h2>
            <span className={`arrow ${isSixthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isSixthPageOpen && (
            <div className="section-content">
              <label className="full-width">
                试验步骤描述 <span className="field-hint">(test_steps)</span>
                <textarea
                  value={formData.test_steps || ""}
                  onChange={(e) => handleChange("test_steps", e.target.value)}
                  rows={12}
                  required
                />
              </label>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsSeventhPageOpen(!isSeventhPageOpen)}
          >
            <h2>附录 E 试验结果</h2>
            <span className={`arrow ${isSeventhPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isSeventhPageOpen && (
            <div className="section-content">
              <div className="full-width table-section">
                <div className="table-header">
                  <h3>E1 试验结果</h3>
                  <button type="button" className="add-btn" onClick={addE1Row}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '50px' }}>序号</th>
                        <th style={{ width: '100px' }}>试验项目</th>
                        <th>测试工况 <span className="field-hint">(test_condition)</span></th>
                        <th>试验评价标准 <span className="field-hint">(evaluation_criteria)</span></th>
                        <th>试验结果 <span className="field-hint">(test_result)</span></th>
                        <th style={{ width: '80px' }}>符合性判定</th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.e1_test_results.map((row, index) => (
                        <tr key={`e1-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td>
                            <textarea 
                              value={row.test_item} 
                              onChange={e => handleE1Change(index, "test_item", e.target.value)} 
                              rows={2}
                            />
                          </td>
                          <td>
                            <textarea 
                              value={row.test_condition} 
                              onChange={e => handleE1Change(index, "test_condition", e.target.value)} 
                              rows={6}
                            />
                          </td>
                          <td>
                            <textarea 
                              value={row.evaluation_criteria} 
                              onChange={e => handleE1Change(index, "evaluation_criteria", e.target.value)} 
                              rows={6}
                            />
                          </td>
                          <td>
                            <textarea 
                              value={row.test_result} 
                              onChange={e => handleE1Change(index, "test_result", e.target.value)} 
                              rows={6}
                            />
                          </td>
                          <td>
                            <input 
                              value={row.compliance} 
                              onChange={e => handleE1Change(index, "compliance", e.target.value)} 
                              style={{ color: row.compliance === "不符合" ? "red" : "inherit" }}
                            />
                          </td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeE1Row(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsEighthPageOpen(!isEighthPageOpen)}
          >
            <h2>附录 E2 试验结果图表</h2>
            <span className={`arrow ${isEighthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isEighthPageOpen && (
            <div className="section-content">
              <div className="upload-btn-wrapper full-width">
                <button type="button" className="btn">+ 批量上传图表照片</button>
                <input type="file" accept="image/*" multiple onChange={handleE2ChartUpload} />
              </div>

              <div className="photo-grid">
                {formData.e2_charts.map((chart, index) => (
                  <div key={chart.id} className="photo-card">
                    <img src={chart.chart_image} alt={`图表 ${chart.id}`} />
                    <input
                      type="text"
                      value={chart.title}
                      onChange={(e) => handleE2ChartTitleChange(index, e.target.value)}
                      placeholder="图表标题（如：图1 测试工况1）"
                    />
                    <button type="button" className="remove-photo-btn" onClick={() => removeE2Chart(index)}>
                      删除
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsNinthPageOpen(!isNinthPageOpen)}
          >
            <h2>附录 E3/E4 试验问题</h2>
            <span className={`arrow ${isNinthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isNinthPageOpen && (
            <div className="section-content">
              <div className="full-width table-section">
                <div className="table-header">
                  <h3>E3 试验问题明细</h3>
                  <button type="button" className="add-btn" onClick={addE3Row}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '50px' }}>序号</th>
                        <th style={{ width: '100px' }}>发生时间</th>
                        <th>问题描述 <span className="field-hint">(description/issue_description)</span></th>
                        <th style={{ width: '80px' }}>问题方</th>
                        <th>解决措施 <span className="field-hint">(solution)</span></th>
                        <th style={{ width: '80px' }}>是否解决</th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.e3_issues.map((row, index) => (
                        <tr key={`e3-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.occur_time} onChange={e => handleE3Change(index, "occur_time", e.target.value)} /></td>
                          <td>
                            <textarea 
                              value={row.description} 
                              onChange={e => handleE3Change(index, "description", e.target.value)} 
                              rows={3}
                            />
                          </td>
                          <td><input value={row.responsible_party} onChange={e => handleE3Change(index, "responsible_party", e.target.value)} /></td>
                          <td>
                            <textarea 
                              value={row.solution} 
                              onChange={e => handleE3Change(index, "solution", e.target.value)} 
                              rows={3}
                            />
                          </td>
                          <td><input value={row.is_resolved} onChange={e => handleE3Change(index, "is_resolved", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeE3Row(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section" style={{ marginTop: '20px' }}>
                <div className="table-header">
                  <h3>E4 试验问题照片</h3>
                </div>
                <div className="upload-btn-wrapper full-width">
                  <button type="button" className="btn">+ 批量上传问题照片</button>
                  <input type="file" accept="image/*" multiple onChange={handleE4PhotoUpload} />
                </div>

                <div className="photo-grid">
                  {formData.e4_issue_photos.map((photo, index) => (
                    <div key={photo.id} className="photo-card">
                      <img src={photo.dataUrl} alt={`问题照片 ${photo.id}`} />
                      <input
                        type="text"
                        value={photo.label}
                        onChange={(e) => handleE4PhotoLabelChange(index, e.target.value)}
                        placeholder="照片描述"
                      />
                      <button type="button" className="remove-photo-btn" onClick={() => removeE4Photo(index)}>
                        删除
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="section">
          <div 
            className="section-header" 
            onClick={() => setIsTenthPageOpen(!isTenthPageOpen)}
          >
            <h2>附录 F 拆解结果</h2>
            <span className={`arrow ${isTenthPageOpen ? 'open' : ''}`}>▼</span>
          </div>
          
          {isTenthPageOpen && (
            <div className="section-content">
              <div className="full-width table-section">
                <div className="table-header">
                  <h3>F1 主要零部件拆解结果</h3>
                  <button type="button" className="add-btn" onClick={addF1Row}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '50px' }}>序号</th>
                        <th>测试零部件名称 <span className="field-hint">(name/part_name)</span></th>
                        <th>零件号 <span className="field-hint">(part_no)</span></th>
                        <th>拆解评价标准 <span className="field-hint">(evaluation_criteria)</span></th>
                        <th>拆解结果 <span className="field-hint">(teardown_result)</span></th>
                        <th style={{ width: '80px' }}>符合性判定</th>
                        <th>备注 <span className="field-hint">(remark)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.f1_teardown_results.map((row, index) => (
                        <tr key={`f1-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.part_name} onChange={e => handleF1Change(index, "part_name", e.target.value)} /></td>
                          <td><input value={row.part_no} onChange={e => handleF1Change(index, "part_no", e.target.value)} /></td>
                          <td><input value={row.evaluation_criteria} onChange={e => handleF1Change(index, "evaluation_criteria", e.target.value)} /></td>
                          <td><input value={row.teardown_result} onChange={e => handleF1Change(index, "teardown_result", e.target.value)} /></td>
                          <td><input value={row.compliance} onChange={e => handleF1Change(index, "compliance", e.target.value)} /></td>
                          <td><input value={row.remark} onChange={e => handleF1Change(index, "remark", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeF1Row(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="full-width table-section" style={{ marginTop: '20px' }}>
                <div className="table-header">
                  <h3>F2 拆解明细</h3>
                  <button type="button" className="add-btn" onClick={addF2Row}>+ 添加行</button>
                </div>
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '50px' }}>序号</th>
                        <th style={{ width: '100px' }}>发生时间</th>
                        <th>问题描述 <span className="field-hint">(description/issue_description)</span></th>
                        <th>描述 <span className="field-hint">(description)</span></th>
                        <th>备注 <span className="field-hint">(remark)</span></th>
                        <th style={{ width: '60px' }}>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.f2_teardown_issues.map((row, index) => (
                        <tr key={`f2-${index}`}>
                          <td className="text-center">{row.id}</td>
                          <td><input value={row.occur_time} onChange={e => handleF2Change(index, "occur_time", e.target.value)} /></td>
                          <td>
                            <textarea 
                              value={row.issue_description} 
                              onChange={e => handleF2Change(index, "issue_description", e.target.value)} 
                              rows={2}
                            />
                          </td>
                          <td>
                            <textarea 
                              value={row.description} 
                              onChange={e => handleF2Change(index, "description", e.target.value)} 
                              rows={2}
                            />
                          </td>
                          <td><input value={row.remark} onChange={e => handleF2Change(index, "remark", e.target.value)} /></td>
                          <td className="text-center">
                            <button type="button" className="del-btn" onClick={() => removeF2Row(index)}>删除</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <div className="full-width table-section" style={{ marginTop: '20px' }}>
                <div className="table-header">
                  <h3>F3 拆解问题照片</h3>
                </div>
                <div className="upload-btn-wrapper full-width">
                  <button type="button" className="btn">+ 批量上传拆解照片</button>
                  <input type="file" accept="image/*" multiple onChange={handleF3PhotoUpload} />
                </div>

                <div className="photo-grid">
                  {formData.f3_teardown_photos.map((photo, index) => (
                    <div key={photo.id} className="photo-card">
                      <img src={photo.dataUrl} alt={`拆解照片 ${photo.id}`} />
                      <input
                        type="text"
                        value={photo.label}
                        onChange={(e) => handleF3PhotoLabelChange(index, e.target.value)}
                        placeholder="照片描述"
                      />
                      <button type="button" className="remove-photo-btn" onClick={() => removeF3Photo(index)}>
                        删除
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {errorMessage ? <div className="error">{errorMessage}</div> : null}

        <button type="submit" disabled={isSubmitting || !isValid}>
          {isSubmitting ? "导出中..." : "导出 Word"}
        </button>
      </form>
    </main>
  );
}

export default App;
