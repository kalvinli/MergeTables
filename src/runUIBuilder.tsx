import { bitable, UIBuilder, IOpenSegmentType, FieldType, SelectOptionsType } from "@lark-base-open/js-sdk";

/*
Text = 1,
Number = 2,
SingleSelect = 3,
MultiSelect = 4,
DateTime = 5,
Checkbox = 7,
User = 11,
Phone = 13,
Url = 15,
Attachment = 17,
SingleLink = 18,
Lookup = 19,
Formula = 20,
DuplexLink = 21,
Location = 22,
GroupChat = 23,
Barcode = 99001,
Progress = 99002,
Currency = 99003,
Rating = 99004
//*/

export default async function main(uiBuilder: UIBuilder) {

  const language: any = {
    zh: {
      user_guide: `> [「多数据表合并」插件使用指南](https://bytespace.feishu.cn/docx/OoEHdlUUYoFJkxxjn47cKVK0nod)`,
      table_Source_List: "源数据表",
      select_view: "源数据表视图名称（不输入或不存在视图则获取全量数据）",
      view_label: "表格",
      table_Target: "目标数据表",
      field_Target: "待合并字段",
      father_field_Target: "父记录字段",
      checkbox_clean_table: "清空目标数据表",
      sync_options_target: "合并单选/多选字段选项到目标表",
      submit_button: "合并数据",
      table_Source_List_notice: "请选择源数据表",
      field_Target_notice: "选择待合并字段",
      clear_data_notice: "正在清除目标表数据，请稍等...",
      prepare_merge_data_notice: "开始准备合并数据，请稍等...",
      finish_notice: "所有数据写入完成",
      check_fields_type_notice: "写入数据格式不正确，请检查源表与目标表字段类型是否一致"
    },
    en: {
      user_guide: "[User Guide](https://bytespace.feishu.cn/docx/OoEHdlUUYoFJkxxjn47cKVK0nod)",
      table_Source_List: "Choose the source tables",
      select_view: "Source datasheet view names (if no input is entered or the view does not exist, the full data will be obtained)",
      view_label: "Grid",
      table_Target: "Choose a target table",
      field_Target: "Choose the fields for merging",
      father_field_Target: "Parent record field",
      checkbox_clean_table: "Clear data of target table",
      sync_options_target: "Merge single/multi-select field options into target table",
      submit_button: "Merge Data",
      table_Source_List_notice: "Please choose the source tables",
      field_Target_notice: "Please choose the fields for merging",
      clear_data_notice: "Clearing data of target table, please wait...",
      prepare_merge_data_notice: "Preparing to merge data, please wait...",
      finish_notice: "All data writing completed",
      check_fields_type_notice: "The format of the written data is incorrect. \nPlease check whether the field types of the source tables and the target table are consistent."
    }
  }

  // 根据本地语言设置lang的值
  let lang: any = await bitable.bridge.getLanguage();
  if (lang === 'zh' || lang === 'zh-TW' || lang === 'zh-HK') {
    lang = 'zh';
  } else {
    lang = 'en';
  }

  const instanceId = await bitable.bridge.getInstanceId();
  const userId = await bitable.bridge.getUserId();
  const getSelection = await bitable.base.getSelection();
  let tableId: any = getSelection.tableId;

  const fieldType_List_All = [1, 2, 3, 4, 5, 7, 11, 13, 15, 17, 18, 19, 20, 21, 22, 23, 99001, 99002, 99003, 99004];
  let options_list: any = [];
  const tables_meta = await bitable.base.getTableMetaList();
  tables_meta.forEach((item: any) => {
    options_list.push({ label: item.name, value: item.id })
  });
  const pre_table = await bitable.base.getTableById(tableId);
  const fields_meta = await pre_table.getFieldMetaList();
  let field_target_list = [];
  for (let field_item of fields_meta) {
    if(fieldType_List_All.indexOf(field_item.type) >= 0 && field_item.type !== 19 && field_item.type !== 20){
      field_target_list.push(field_item.id);
    }
  }

  let mergetables_local: any = localStorage.getItem("mergetables_" + instanceId + "_" + userId + "_" + tableId);
  if (!mergetables_local) {
    mergetables_local = { "table_Source_List": [], "select_view": [], "table_Target": "", "father_field_Target": [], "checkbox_clean_table": [], "sync_options_target": [], "field_Target": field_target_list };
  } else {
    mergetables_local = JSON.parse(mergetables_local);
  }
  

  uiBuilder.markdown(language[lang].user_guide);
  uiBuilder.form((form: any) => ({
    formItems: [
      form.select('table_Source_List', {
        label: language[lang].table_Source_List,
        options: options_list,
        multiple: true,
        defaultValue: mergetables_local.table_Source_List
      }),
      form.select('select_view', {
        label: language[lang].select_view,
        options: [{ label: language[lang].view_label, value: language[lang].view_label }],
        defaultValue: mergetables_local.select_view,
        tags: true,
      }),
      form.tableSelect('table_Target', { label: language[lang].table_Target }),
      form.fieldSelect('field_Target', {
        label: language[lang].field_Target,
        sourceTable: 'table_Target',
        multiple: true,
        filter: ({ type }: { type: any }) => (fieldType_List_All.indexOf(type) >= 0 && type !== 19 && type !== 20),
        optionFilterProp: "children",
        showSearch: true,
        filterOption: (input: string, option: any) => `${(option?.label ?? '')}`.toLowerCase().includes(input.toLowerCase()),
        defaultValue: mergetables_local.field_Target
      }),
      form.fieldSelect('father_field_Target', {
        label: language[lang].father_field_Target,
        sourceTable: 'table_Target',
        multiple: true,
        filter: ({ type }: { type: any }) => (type === 18 || type === 21),
        optionFilterProp: "children",
        showSearch: true,
        filterOption: (input: string, option: any) => `${(option?.label ?? '')}`.toLowerCase().includes(input.toLowerCase()),
        defaultValue: mergetables_local.father_field_Target
      }),
      form.checkboxGroup('checkbox_clean_table', {
        label: '',
        options: [language[lang].checkbox_clean_table],
        defaultValue: mergetables_local.checkbox_clean_table
      }),
      form.checkboxGroup('sync_options_target', {
        label: '',
        options: [language[lang].sync_options_target],
        defaultValue: mergetables_local.sync_options_target
      }),

    ],
    buttons: [language[lang].submit_button],

  }), async ({ values }: { values: any }) => {
    let { table_Source_List, select_view, table_Target, field_Target, father_field_Target, checkbox_clean_table, sync_options_target } = values;
    // console.log(values);

    // 分组函数
    function grouping(array: any, subGroupLength: any) {
      let index = 0;
      let newArray = [];
      while (index < array.length) {
        newArray.push(array.slice(index, index += subGroupLength));
      }
      return newArray;
    }

    // 判断以下字段是否进行了选择，如果未选择则提示并返回
    if (table_Source_List.length === 0) { alert(language[lang].table_Source_List_notice); return; }
    if (field_Target.length === 0) { alert(language[lang].field_Target_notice); return; }

    let field_target_tmp: any = [];
    let values_tmp: any = {};
    let field_item: any;
    for (field_item of field_Target) {
      field_target_tmp.push(field_item.id);
    }

    values_tmp.table_Source_List = table_Source_List;
    values_tmp.select_view = select_view;
    values_tmp.table_Target = table_Target.id;
    values_tmp.checkbox_clean_table = checkbox_clean_table;
    values_tmp.sync_options_target = sync_options_target;
    values_tmp.field_Target = field_target_tmp;
    if (typeof father_field_Target === 'undefined') father_field_Target = [];
    if (father_field_Target.length !== 0) {
      values_tmp.father_field_Target = [father_field_Target[0]?.id];
    } else {
      values_tmp.father_field_Target = [];
    }


    // console.log(values_tmp);

    localStorage.setItem("mergetables_" + instanceId + "_" + userId + "_" + tableId, JSON.stringify(values_tmp));

    if (checkbox_clean_table.length > 0) {
      const recordIdList = await table_Target.getRecordIdList();
      if (recordIdList.length > 0) {
        var confirm_msg: any = false;
        if (lang === "zh") {
          var msg = "\n    请确认是否要清空目标数据表中的 " + String(recordIdList.length) + " 条数据？\n";
        } else {
          if (recordIdList.length === 1) {
            var msg = "\n    Please confirm whether you will to clear 1 piece of data in the target table?\n";
          } else {
            var msg = "\n    Please confirm whether you will to clear " + String(recordIdList.length) + " pieces of data in the target table?？\n";
          }
        }
        confirm_msg = confirm(msg);

        if (confirm_msg == true) {
          uiBuilder.showLoading(language[lang].clear_data_notice);
          const new_recordIdList = grouping(recordIdList, 5000);
          for (let i = 0; i < new_recordIdList.length; i++) {
            await table_Target.deleteRecords(new_recordIdList[i]);
            // // 延迟代码
            // if (i < new_recordIdList.length - 1) {
            //   await new Promise((resolve) => {
            //     setTimeout(() => {
            //       resolve("finished");
            //     }, 3000);
            //   })
            // }
          }
        } else { return; }
      }
    }

    uiBuilder.showLoading(language[lang].prepare_merge_data_notice);

    // 获取数据表的名称
    const table_target = table_Target;

    // 根据选择的字段信息生成包含字段id和name的数组
    let field_name_list: any = [];
    const metaList_target: any = await table_target.getFieldMetaList();
    metaList_target.forEach((target_item: any) => {
      field_Target.forEach((field_item: any) => {
        if (target_item.id == field_item.id) {
          field_name_list.push({ field_id: target_item.id, name: target_item.name, type: target_item.type });
        }
      })
    })

    // console.log(field_name_list);

    // console.log(table_Source_List);
    let merge_field_name_list: any = [];
    let records_update_list: any = [];
    let record_update_list: any = {};
    let fields_update_list: any = { 'fields': {} };
    let select_property: any = {};
    let option_list = [];

    // 循环处理多个源表的数据
    for (const table_Source of table_Source_List) {

      //根据前面的field_name_list重新生成包含name,type和源和目标数据表field_id的数组
      const table_source = await bitable.base.getTableById(table_Source as string);
      const table_source_name = await table_source.getName();
      const metaList_source: any = await table_source.getFieldMetaList();
      merge_field_name_list = [];
      // console.log(metaList_source);
      metaList_source.forEach((source_item: any) => {
        field_name_list.forEach((target_item: any) => {
          if (source_item.name === target_item.name) {
            merge_field_name_list.push({ name: target_item.name, type: target_item.type, field_id: { source_id: source_item.id, target_id: target_item.field_id } })
          }
        })
      })
      // console.log(1, merge_field_name_list);

      let hasMore: boolean = true;
      let pageSize: number = 5000;
      let pageToken: string = "";
      let viewId: string = "";
      let count: number = 0;
      const merge_field_name_list_len = merge_field_name_list.length;
      while (hasMore) {
        // console.log(1, dataindex);

        if (typeof select_view !== 'undefined') {
          if (select_view.length > 0 && select_view[0] !== '') {
            const viewMetaList = await table_source.getViewMetaList();
            viewMetaList.forEach((item: any) => {
              if (item.name === select_view[0]) {
                viewId = item.id;
              }
            })
          }
        }

        const source_recordValueList = await table_source.getRecords({ pageSize: pageSize, pageToken: pageToken, viewId: viewId });
        // console.log(2, source_recordValueList);
        pageToken = source_recordValueList.pageToken || '';
        hasMore = source_recordValueList.hasMore;
        const get_records = source_recordValueList.records;
        const recordid_list_sourcre_len = source_recordValueList.total;
        // console.log(get_records);

        // 循环处理字段数组
        for (var i = 0; i < get_records.length; i++) {
          const record_items: any = get_records[i];
          // console.log(record_items);
          for (var j = 0; j < merge_field_name_list_len; j++) {
            let record_value: any = '';
            const record_item: any = record_items.fields[merge_field_name_list[j].field_id.source_id];
            const merge_field_target_id = merge_field_name_list[j].field_id.target_id;

            // console.log(111, record_item);

            // console.log(merge_field_name_list[j].type);

            let field_item: any = "";
            let new_field: any = "";
            let field_name = '源表父记录ID';
            if (lang !== 'zh') {
              field_name = 'Source Parent ID';
            }

            switch (merge_field_name_list[j].type) {

              case 3: //SingleSelect
                record_value = record_item?.text || '';
                // 获取目标表单选字段的选项信息
                let get_ss_options: any = "";
                // if (table_source_pre === table_Source) {

                if (typeof select_property[merge_field_target_id] === 'undefined') {
                  const target_ss_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                  select_property[merge_field_target_id] = target_ss_record_item.property;
                  get_ss_options = target_ss_record_item?.property?.options || '';
                } else {
                  get_ss_options = select_property[merge_field_target_id].options;
                }

                if (sync_options_target.length > 0) {
                  if (get_ss_options.length === 0 && record_value !== "") {
                    await table_target.setField(
                      merge_field_target_id,
                      {
                        type: FieldType.SingleSelect,
                        property: {
                          options: [{ name: record_value }],
                        },
                        optionsType: SelectOptionsType.STATIC,
                      }
                    );
                    const target_ss_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                    select_property[merge_field_target_id] = target_ss_record_item.property;
                    get_ss_options = target_ss_record_item?.property?.options || '';
                  }
                }

                for (var k = 0; k < get_ss_options.length; k++) {
                  if (record_value == get_ss_options[k].name) {
                    record_value = get_ss_options[k];
                    break;
                  }

                  if (k === get_ss_options.length - 1 && record_value !== "") {
                    if (sync_options_target.length > 0) {
                      try {
                        get_ss_options.push({ name: record_value });
                        await table_target.setField(
                          merge_field_target_id,
                          {
                            type: FieldType.SingleSelect,
                            property: {
                              options: get_ss_options,
                            },
                            optionsType: SelectOptionsType.STATIC,
                          }
                        );
                      } catch (e) {
                        //
                      }
                      const target_ss_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                      select_property[merge_field_target_id] = target_ss_record_item.property;
                      get_ss_options = target_ss_record_item?.property?.options || '';
                      record_value = get_ss_options[get_ss_options.length - 1];
                    }
                  }
                }
                break;

              case 4: //MultiSelect
                record_value = record_item || '';
                // 获取目标表多选字段的选项信息
                let get_ms_options: any = "";
                if (typeof select_property[merge_field_target_id] === 'undefined') {
                  const target_ms_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                  select_property[merge_field_target_id] = target_ms_record_item.property;
                  get_ms_options = target_ms_record_item?.property?.options || '';
                } else {
                  get_ms_options = select_property[merge_field_target_id].options;
                }

                if (sync_options_target.length > 0) {
                  if (get_ms_options.length === 0 && record_value !== "") {
                    option_list = [];
                    for (var l = 0; l < record_value.length; l++) {
                      option_list.push({ name: record_value[l].text })
                    }
                    await table_target.setField(
                      merge_field_target_id,
                      {
                        type: FieldType.MultiSelect,
                        property: {
                          options: option_list,
                        },
                        optionsType: SelectOptionsType.STATIC,
                      }
                    );
                    const target_ms_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                    select_property[merge_field_target_id] = target_ms_record_item.property;
                    get_ms_options = target_ms_record_item?.property?.options || '';
                  }
                }

                let get_ms_options_value: any = [];
                for (var l = 0; l < record_value.length; l++) {
                  for (var k = 0; k < get_ms_options.length; k++) {
                    if (record_value[l].text == get_ms_options[k].name) {
                      get_ms_options_value.push(get_ms_options[k]);
                      break;
                    }
                    if (k === get_ms_options.length - 1 && record_value !== "") {
                      if (sync_options_target.length > 0) {
                        try {
                          get_ms_options.push({ name: record_value[l].text });
                          await table_target.setField(
                            merge_field_target_id,
                            {
                              type: FieldType.MultiSelect,
                              property: {
                                options: get_ms_options,
                              },
                              optionsType: SelectOptionsType.STATIC,
                            }
                          );
                        } catch (e) {
                          //
                        }
                        const target_ms_record_item = await table_target.getFieldMetaById(merge_field_target_id);
                        select_property[merge_field_target_id] = target_ms_record_item.property;
                        get_ms_options = target_ms_record_item?.property?.options || '';
                        get_ms_options_value.push(get_ms_options[get_ms_options.length - 1]);
                        break;
                      }
                    }
                  }
                }
                record_value = get_ms_options_value;
                break;

              case 7: //Checkbox
                record_value = record_item ? true : false;
                break;

              case 2: //Number
              case 99002: //Progress
              case 99003: //Currency
              case 99004: //Rating
                if (record_item === null) {
                  record_value = null;
                } else {
                  if (typeof record_item === "string") {
                    try {
                      record_value = parseFloat(record_item);
                    } catch (e) {
                      record_value = null;
                    }
                  } else {
                    record_value = record_item[0];
                  }
                }
                if (typeof record_value === 'undefined') {
                  record_value = record_item;
                }
                break;

              case 1: //Text
              case 5: //DateTime
              case 11: //User
              case 13: //Phone
              case 15: //Url
              case 17: //Attachment
              case 18: //SingleLink
              case 21: //DuplexLink
              case 22: //Location
              case 23: //GroupChat
              case 99001: //Barcode

                if (father_field_Target.length !== 0 && merge_field_target_id === father_field_Target[0]?.id) {
                  try {  // 获取字段
                    field_item = await table_Target.getFieldByName(field_name);
                    new_field = field_item.id
                  } catch (e) {  // 获取出错后添加字段
                    new_field = await table_Target.addField({
                      type: FieldType.Text,
                      property: null,
                      name: field_name,
                    });
                  }
                  try {
                    record_value = record_item.recordIds[0];
                  } catch (e) {
                    record_value = null;
                  }
                } else {
                  record_value = record_item;
                }
                break;

              default:
                record_value = null;
                break;
            }
            if (new_field !== "") {
              record_update_list[new_field] = record_value;
            } else {
              record_update_list[merge_field_target_id] = record_value;
            }
          }
          if (lang === 'zh') {
            uiBuilder.showLoading(`正在处理【` + table_source_name + `】表的第 ` + String(count + 1) + ` / ` + recordid_list_sourcre_len + ` 条记录`);
          }
          count++;
          fields_update_list.fields = record_update_list;
          // record_update_list = {}; // 写入空记录
          records_update_list.push({ "fields": record_update_list });
          record_update_list = {};
          fields_update_list = {};
        }
      }
    }
    // console.log(records_update_list);
    let new_records_update_list = grouping(records_update_list, 5000);

    // console.log(new_records_update_list);
    const date1: any = new Date();
    for (let ii = 0; ii < new_records_update_list.length; ii++) {
      if (lang === 'zh') {
        uiBuilder.showLoading(`正在写入第 ` + String(ii + 1) + ` / ` + new_records_update_list.length + ` 页的 ` + new_records_update_list[ii].length + ` 条记录`);
      }
      const date2: any = new Date();
      try {
        await table_target.addRecords(new_records_update_list[ii]);
      } catch (e) {
        alert(language[lang].check_fields_type_notice);
        uiBuilder.hideLoading();
        return;
      }

      // // 延迟代码
      // if (ii < new_records_update_list.length - 1) {
      //   await new Promise((resolve) => {
      //     setTimeout(() => {
      //       resolve("finished");
      //     }, 3000);
      //   })
      // }

      const date3: any = new Date();
      console.log("第" + String(ii + 1) + "次写入时长：", String((date3 - date2) / 1000));
    }
    const date4: any = new Date();
    console.log("总写入时长：", String((date4 - date1) / 1000));

    // 隐藏加载提示
    uiBuilder.hideLoading();
    uiBuilder.message.success(language[lang].finish_notice);

  });
}