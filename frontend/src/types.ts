export type AssemblyInfo = {
  id: number;
  name: string;
  part_no: string;
  version: string;
  batch_no: string;
  remark: string;
};

export type ComponentInfo = {
  id: number;
  name: string;
  part_no: string;
  version: string;
  trace_no: string;
  supplier: string;
  remark: string;
};

export type PhotoInfo = {
  id: number;
  label: string;
  dataUrl: string; // Base64 image data
};

export type TestTimeInfo = {
  id: number;
  stage: string;
  time: string;
  env_temp: string;
  env_humidity: string;
  location: string;
};

export type TestEquipmentInfo = {
  id: number;
  name: string;
  equip_no: string;
  quantity: string;
  valid_date: string;
  remark_image: string; // Base64 image
};

export type TestResultInfo = {
  id: number;
  test_item: string;
  test_condition: string;
  evaluation_criteria: string;
  test_result: string;
  compliance: string;
};

export type ChartInfo = {
  id: number;
  title: string;
  chart_image: string; // Base64 image
};

export type IssueInfo = {
  id: number;
  occur_time: string;
  description: string;
  responsible_party: string;
  solution: string;
  is_resolved: string;
};

export type TeardownResultInfo = {
  id: number;
  part_name: string;
  part_no: string;
  evaluation_criteria: string;
  teardown_result: string;
  compliance: string;
  remark: string;
};

export type TeardownIssueInfo = {
  id: number;
  occur_time: string;
  issue_description: string;
  description: string;
  remark: string;
};

export type ReportFormData = {
  // 第一页字段
  report_no: string;
  project_name: string;
  report_title: string;
  subject: string;
  prepared_by: string;
  prepared_date: string;
  reviewed_by: string;
  reviewed_date: string;
  approved_by: string;
  approved_date: string;
  company_name: string;

  // 第二页字段
  sample_name: string;
  sample_project: string;
  production_stage: string;
  sample_quantity: string;
  sample_no: string;
  test_location: string;
  send_date: string;
  test_date: string;
  task_source: string;
  test_item: string;
  test_basis_method: string;
  test_basis_criteria: string;
  test_conclusion: string;
  test_attachments: string;

  // 第三页（附录）字段
  test_purpose: string;
  assembly_list: AssemblyInfo[];
  component_list: ComponentInfo[];
  photo_list: PhotoInfo[];

  // 附录 B4 电机参数表 (表B4-1)
  motor_manufacturer: string;
  motor_model: string;
  motor_type: string;
  motor_rated_efficiency: string;
  motor_peak_power: string;
  motor_rated_power: string;
  motor_max_torque: string;
  motor_rated_torque: string;
  motor_max_speed: string;
  motor_rated_speed: string;
  motor_cooling_method: string;
  motor_water_pipe_dia: string;
  motor_water_inlet_temp: string;
  motor_water_flow: string;
  motor_test_state: string;
  motor_steel_stamp: string;

  // 附录 B4 控制器参数表 (表B4-2)
  ctrl_manufacturer: string;
  ctrl_model: string;
  ctrl_continuous_current: string;
  ctrl_work_system: string;
  ctrl_short_current: string;
  ctrl_phase_num: string;
  ctrl_cooling_method: string;
  ctrl_factory_no: string;
  ctrl_software_program: string;

  // 附录 B4 整桥参数表 (表B4-3)
  bridge_manufacturer: string;
  bridge_model: string;
  bridge_voltage: string;
  bridge_max_output_torque: string;
  bridge_rated_power: string;
  bridge_peak_power: string;
  bridge_cooling_method: string;
  bridge_gear_ratio: string;
  bridge_max_output_speed: string;
  bridge_rated_load: string;

  // 附录 C 字段
  test_leader: string;
  test_engineer: string;
  c_cooling_method: string;
  c_cooling_pressure_flow: string;
  c_cooling_water_temp: string;

  // 附录 C2 试验时间表格
  test_time_list: TestTimeInfo[];

  // 附录 C3 主要设备表格
  equipment_list: TestEquipmentInfo[];

  // 附录 D 试验步骤
  test_steps: string;

  // 附录 E1 试验结果表格
  e1_test_results: TestResultInfo[];

  // 附录 E2 试验结果图表
  e2_charts: ChartInfo[];

  // 附录 E3 试验问题明细
  e3_issues: IssueInfo[];

  // 附录 E4 试验问题照片
  e4_issue_photos: PhotoInfo[];

  // 附录 F1 拆解结果表格
  f1_teardown_results: TeardownResultInfo[];

  // 附录 F2 拆解明细表格
  f2_teardown_issues: TeardownIssueInfo[];

  // 附录 F3 拆解问题照片
  f3_teardown_photos: PhotoInfo[];
};
