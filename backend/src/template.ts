import fs from "node:fs";
import path from "node:path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
// @ts-ignore
import ImageModule from "docxtemplater-image-module-free";
import { ReportFormData } from "./types";

const requiredFields: Array<keyof ReportFormData> = [
  "report_no",
  "project_name",
  "report_title",
  "subject",
  "prepared_by",
  "prepared_date",
  "reviewed_by",
  "reviewed_date",
  "approved_by",
  "approved_date",
  "company_name"
];

export function validateFormData(input: any): ReportFormData {
  if (typeof input !== "object" || input === null) {
    throw new Error("请求数据格式错误");
  }

  // 把所有的 undefined 或 null 转换为 空字符串，防止前端漏传字段时模板引擎报错或者显示 undefined
  const safeData: any = {};
  for (const key in input) {
    if (Object.prototype.hasOwnProperty.call(input, key)) {
      if (Array.isArray(input[key])) {
        safeData[key] = input[key];
      } else {
        safeData[key] = input[key] == null ? "" : String(input[key]);
      }
    }
  }

  const normalizedData: Partial<ReportFormData> = {};

  for (const field of requiredFields) {
    const value = safeData[field];
    if (typeof value !== "string" || value.trim().length === 0) {
      throw new Error(`字段 ${field} 不能为空`);
    }
    (normalizedData as any)[field] = value.trim();
  }

  return { ...safeData, ...normalizedData } as ReportFormData;
}

export function generateWordBuffer(data: ReportFormData): Buffer {
  const configuredPath = process.env.TEMPLATE_PATH;
  const templatePath =
    configuredPath && configuredPath.trim().length > 0
      ? configuredPath
      : path.resolve(process.cwd(), "templates/test_word.docx");

  if (!fs.existsSync(templatePath)) {
    throw new Error(`未找到模板文件: ${templatePath}`);
  }

  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);

  // 设置图片模块
  const imageOptions = {
    getImage: function (tagValue: string) {
      if (!tagValue) return Buffer.from("");
      // tagValue 会是 base64 字符串 "data:image/png;base64,iVBOR..."
      return Buffer.from(tagValue.replace(/^data:image\/\w+;base64,/, ""), "base64");
    },
    getSize: function (_img: any, _tagValue: string, tagName: string) {
      // 区分照片列表和设备列表的图片大小
      if (tagName === "remark_image") {
        return [100, 50]; // 设备列表中的图片尺寸较小
      }
      if (tagName === "chart_image") {
        return [600, 350]; // 图表尺寸，稍微宽一些
      }
      return [500, 350]; // 默认大图尺寸
    }
  };
  const imageModule = new ImageModule(imageOptions);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '[[', end: ']]' },
    modules: [imageModule]
  });

  const equipmentList = (data.equipment_list || []).map((item) => ({
    ...item,
    has_image: !!item.remark_image,
  }));

  doc.render({
    ...data,
    equipment_list: equipmentList,
    // 兼容原模板中带尖括号和星号的特殊标记
    "**编制签字**": data.prepared_by,
    "**编制签字日期**": data.prepared_date,
    "**审核签字**": data.reviewed_by,
    "**审核签字日期**": data.reviewed_date,
    "**批准签字**": data.approved_by,
    "**批准签字日期**": data.approved_date,
    // 增加一个替换逻辑，避免原模板的硬编码
    "EA3800N-SYBG-202505-001": data.report_no,
    "EA3800N项目": data.project_name,
    "EA3800N整桥温升试验报告（第三轮）": data.subject,
    
    // 兼容第二页的内容（如果没有被替换成[[]]占位符）
    "EA3800N 整桥": data.sample_name,
    "EA3800N": data.sample_project,
    "B样": data.production_stage,
    "1": data.sample_quantity,
    "EA3800N101-202500417-001": data.sample_no,
    "杭州时代电动试验室": data.test_location,
    "2025/4/17": data.send_date,
    "2025/4/28": data.test_date,
    "试验申请单": data.task_source,
    "整桥温升试验": data.test_item,
    "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）": data.test_basis_criteria
  });
  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE"
  });
}

export function generateWordBufferCustom(data: ReportFormData, customTemplatePath: string): Buffer {
  if (!fs.existsSync(customTemplatePath)) {
    throw new Error(`未找到模板文件: ${customTemplatePath}`);
  }

  const content = fs.readFileSync(customTemplatePath, "binary");
  const zip = new PizZip(content);

  const imageOptions = {
    getImage: function (tagValue: string) {
      return Buffer.from(tagValue.replace(/^data:image\/\w+;base64,/, ""), "base64");
    },
    getSize: function () {
      return [500, 350];
    }
  };
  const imageModule = new ImageModule(imageOptions);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '[[', end: ']]' },
    modules: [imageModule]
  });

  doc.render({
    ...data,
    // 兼容原模板中带尖括号和星号的特殊标记
    "**编制签字**": data.prepared_by,
    "**编制签字日期**": data.prepared_date,
    "**审核签字**": data.reviewed_by,
    "**审核签字日期**": data.reviewed_date,
    "**批准签字**": data.approved_by,
    "**批准签字日期**": data.approved_date,
    // 增加一个替换逻辑，避免原模板的硬编码
    "EA3800N-SYBG-202505-001": data.report_no,
    "EA3800N项目": data.project_name,
    "EA3800N整桥温升试验报告（第三轮）": data.subject,
    
    // 兼容第二页的内容（如果没有被替换成[[]]占位符）
    "EA3800N 整桥": data.sample_name,
    "EA3800N": data.sample_project,
    "B样": data.production_stage,
    "1": data.sample_quantity,
    "EA3800N101-202500417-001": data.sample_no,
    "杭州时代电动试验室": data.test_location,
    "2025/4/17": data.send_date,
    "2025/4/28": data.test_date,
    "试验申请单": data.task_source,
    "整桥温升试验": data.test_item,
    "在2h内稳定在最高许用温度以下的某个温度0.5h以上，或不高于最高许用温度，监测电机最高温度点（电机温度≤140℃，油温≤120℃）": data.test_basis_criteria,

    // B4 表格硬编码替换
    "龙芯电驱动": data.motor_manufacturer,
    "TZ220XYC01": data.motor_model,
    "永磁同步电机": data.motor_type,
    "≥96%": data.motor_rated_efficiency,
    "300": data.motor_peak_power,
    "200": data.motor_rated_power,
    "460": data.motor_max_torque,
    "13000": data.motor_max_speed,
    "9550": data.motor_rated_speed,
    "油冷": data.motor_cooling_method,
    "15": data.motor_water_pipe_dia,
    "25±2": data.motor_water_inlet_temp,
    "≥15L/min": data.motor_water_flow,
    "批量件": data.motor_test_state,

    "法拉第科技有限公司": data.ctrl_manufacturer,
    "FED200-54Z232S-HZSD6": data.ctrl_model,
    "246A(AC)": data.ctrl_continuous_current,
    "S9": data.ctrl_work_system,
    "550A": data.ctrl_short_current,
    "三相": data.ctrl_phase_num,
    "水冷": data.ctrl_cooling_method,
    "59120334G1020004": data.ctrl_factory_no,
    "FED20_TM_B1.08M184S1V0.02_250201_Encrypt3_250601_Encrypt3.mot16\nFED20_ISG_B1.08M184S1V0.02_250201_Encrypt3_250601_Encrypt3.mot16": data.ctrl_software_program,

    "CDTL": data.bridge_manufacturer,
    "630V": data.bridge_voltage,
    "38511Nm": data.bridge_max_output_torque,
    "200kw": data.bridge_rated_power,
    "300kw": data.bridge_peak_power,
    "系统内油冷 系统外水冷": data.bridge_cooling_method,
    "83.72/21.4": data.bridge_gear_ratio,
    "607": data.bridge_max_output_speed,
    "13T": data.bridge_rated_load
  });
  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE"
  });
}
