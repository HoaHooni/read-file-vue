<template>
  <div class="hello">
    <h1>{{ msg }}</h1>
    <h3>Create XLSX</h3>
    <div>
      <input v-model="sheetName" placeholder="type a new sheet name" />
      <button v-if="sheetName" @click="addSheet">Add Sheet</button>
    </div>
    <div>Sheets: {{ sheets }}</div>
    <xlsx-workbook>
      <xlsx-sheet
        :collection="sheet.data"
        v-for="sheet in sheets"
        :key="sheet.name"
        :sheet-name="sheet.name"
      />
      <xlsx-download>
        <button>Download</button>
      </xlsx-download>
    </xlsx-workbook>
    <hr />
    <section>
      <h3>Import XLSX</h3>
      <input type="file" @change="onChange" />
      <XlsxWorkbook @change="onChange" @created="onCreated" />
      <xlsx-read :file="file">
        <xlsx-sheets>
          <template #default="{ sheets }">
            <select v-model="selectedSheet">
              <option v-for="sheet in sheets" :key="sheet" :value="sheet">
                {{ sheet }}
              </option>
            </select>
          </template>
        </xlsx-sheets>
        <xlsx-table :sheet="selectedSheet" />
        <xlsx-json :sheet="selectedSheet">
          <template #default="{ collection }">
            <div>
              {{ collection }}
            </div>
          </template>
        </xlsx-json>
      </xlsx-read>
    </section>
  </div>
</template>

<script>
import {
  XlsxRead,
  XlsxTable,
  XlsxSheets,
  XlsxJson,
  XlsxWorkbook,
  XlsxSheet,
  XlsxDownload,
} from "vue3-xlsx";
import XLSX from "xlsx";
export default {
  components: {
    XlsxRead,
    XlsxTable,
    XlsxSheets,
    XlsxJson,
    XlsxWorkbook,
    XlsxSheet,
    XlsxDownload,
  },
  name: "HelloWorld",
  props: {
    msg: String,
  },
  data() {
    return {
      file: null,
      selectedSheet: null,
      sheetName: null,
      sheets: [{ name: "SheetOne", data: [{ c: 2 }] }],
      collection: [{ a: 1, b: 2 }],
    };
  },
  methods: {
    onChange(event) {
      var reader = new FileReader();
      reader.onload = function (e) {
        console.log("reader", e);
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        console.log("workbook", workbook);
        let sheetName = workbook.SheetNames[0];

        let worksheet = workbook.Sheets[sheetName];
        console.log("worksheet", worksheet);
        console.log("sheet_to_json", XLSX.utils.sheet_to_json(worksheet));
        console.log("sheet_to_row_object_array", XLSX.utils.sheet_to_json(worksheet, {header : 1}))
      };
      console.log("reader", reader);

      console.log(event);
      this.file = event.target.files ? event.target.files[0] : null;
      reader.readAsArrayBuffer(this.file);
      console.log(this.file);
    },
    addSheet() {
      this.sheets.push({ name: this.sheetName, data: [...this.collection] });
      this.sheetName = null;
    },
  },
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
