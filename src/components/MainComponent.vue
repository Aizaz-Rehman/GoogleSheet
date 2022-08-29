<template>
  <div>
    <v-container class="mt-5">
      <v-row justify="center">
        <h4>Generate Keywords</h4>
      </v-row>
    </v-container>
    <v-container>
      <v-row justify="center" no-gutters>
        <v-col lg="6">
          <v-file-input
            label="Click here to import file"
            outlined
            prepend-icon="mdi-file"
            v-model="selectXlsx"
            dense
            show-size
          ></v-file-input>
        </v-col>
        <v-col lg="2">
          <v-btn color="success" class="ms-2" @click="generateKeywords()">
            Process Data
          </v-btn>
        </v-col>
      </v-row>
      <v-row justify="center"> </v-row>
    </v-container>
    <v-container class="mt-5">
      <v-row justify="center">
        <h4>Generate Master Keywords</h4>
      </v-row>
    </v-container>
    <v-container>
      <v-row justify="center" no-gutters>
        <v-col lg="6">
          <v-file-input
            label="Click here to import file"
            outlined
            prepend-icon="mdi-file"
            v-model="selectSheet"
            dense
            show-size
          ></v-file-input>
        </v-col>
        <v-col lg="2">
          <v-btn color="success" class="ms-2" @click="generateMasterKeywords()">
            Process Data
          </v-btn>
        </v-col>
      </v-row>
      <v-row justify="center"> </v-row>
      <!-- <clip-loader v-if="loading" :color="color1" :size="size"></clip-loader> -->
    </v-container>
  </div>
</template>

<script>
import XLSX from "xlsx";
import toastr from "toastr";
// import ClipLoader from "vue-spinner/src/ClipLoader.vue";
export default {
  name: "App",

  data() {
    return {
      color1: "#0D47A1",
      size: "50px",
      selectXlsx: null,
      selectSheet: null,
      Keywords: [],
      SplittedWords: [],
      ResultKeywords: [],
      MasterKwrds: [],
      RootKeywords: [],
      SplittedRootWords: [],
      ResultRootKeywords: [],
      loading: false,
      alert: false,
      err: null,
      popup: false,
      totalKwrds: null,
    };
  },
  methods: {
    generateKeywords() {
      if (!this.selectXlsx) {
        toastr.options = {
          closeButton: true,
          debug: false,
          positionClass: "toast-top-center",
          onclick: null,
          showDuration: "300",
          hideDuration: "1000",
          timeOut: "3000",
          extendedTimeOut: "1000",
          showEasing: "swing",
          hideEasing: "linear",
          showMethod: "fadeIn",
          hideMethod: "fadeOut",
          opacity: "100",
          Heading: [],
        };
        toastr.error('Please upload a xlsx file"');
        return;
      }
      if (this.selectXlsx) {
        this.loading = true;
        const reader = new FileReader();
        reader.onload = (e) => {
          /* Parse data */
          try {
            const bstr = e.target.result;
            const wb = XLSX.read(bstr, { type: "binary" });
            /* Get first worksheet */
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            this.Heading = data[0];
            for (var i = 1; i < data.length; i++) {
              this.Keywords.push(data[i][0]);
            }

            this.Keywords.forEach((element) => {
              this.SplittedWords.push(element.trim().split(/\s+/));
            });

            this.ResultKeywords = this.getSpilltedWords(this.SplittedWords);

            this.ResultKeywords = this.dupCounts(this.ResultKeywords);
            this.ResultKeywords = this.ResultKeywords.sort(
              (a, b) => b[1] - a[1]
            );
            var singelKeywords = XLSX.utils.json_to_sheet(this.ResultKeywords, {
              skipHeader: true,
            });
            let Heading = [["Keywords", "Count", "Percentage"]];
            XLSX.utils.sheet_add_aoa(singelKeywords, Heading);
            XLSX.utils.sheet_add_json(singelKeywords, this.ResultKeywords, { origin: 'A2', skipHeader: true });
            XLSX.utils.book_append_sheet(wb, singelKeywords, "Single Words");
            this.loading = false;
            XLSX.writeFile(wb, "Analyzed Keywords Sheet.xlsx");
          } catch (error) {
            this.popup = true;
            this.err = error;
          }
        };
     
        reader.readAsBinaryString(this.selectXlsx);
      }
    },
    generateMasterKeywords() {
      if (!this.selectSheet) {
        toastr.options = {
          closeButton: true,
          debug: false,
          positionClass: "toast-top-center",
          onclick: null,
          showDuration: "300",
          hideDuration: "1000",
          timeOut: "3000",
          extendedTimeOut: "1000",
          showEasing: "swing",
          hideEasing: "linear",
          showMethod: "fadeIn",
          hideMethod: "fadeOut",
          opacity: "100",
        };
        toastr.error('Please upload a xlsx file"');
        return;
      }
      if (this.selectSheet) {
        this.loading = true;

        const myReader = new FileReader();

        myReader.onload = (e) => {
          /* Parse data */

          try {
            const bstr = e.target.result;
            const wb = XLSX.read(bstr, { type: "binary" });
            /* Get first worksheet */
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            const nkeywsname = wb.SheetNames[2];
            const nkwordsws = wb.Sheets[nkeywsname];
            var neverKwrds = XLSX.utils.sheet_to_json(nkwordsws, {
              header: 1,
            });
            for (var i = 1; i < data.length; i++) {
              this.Keywords.push(data[i][0]);
            }
            //filtering empty items
            this.Keywords = this.Keywords.filter((e) => e != "");
            neverKwrds = neverKwrds.filter((e) => e != "");
            for (var o = 1; o < data.length; o++) {
              var found = false;
              neverKwrds.forEach((neverword) => {
                if (String(data[o][0]).includes(String(neverword))) {
                  found = true;
                }
              });
              if (found) continue;
              this.MasterKwrds.push(data[o]);
            }
            this.Keywords.forEach((element) => {
              this.SplittedWords.push(element.trim().split(/\s+/));
            });
            this.ResultKeywords = this.getSpilltedWords(this.SplittedWords);
            this.ResultKeywords = this.dupCounts(this.ResultKeywords);
            for (var q = 0; q < this.MasterKwrds.length; q++) {
              this.RootKeywords.push(this.MasterKwrds[q][0]);
            }
            this.RootKeywords.forEach((masterword) => {
              this.SplittedRootWords.push(masterword.trim().split(/\s+/));
            });
            this.ResultRootKeywords = this.getSpilltedWords(
              this.SplittedRootWords
            );
            this.ResultRootKeywords = this.dupCounts(this.ResultRootKeywords);

            this.ResultRootKeywords = this.ResultRootKeywords.sort(
              (a, b) => b[1] - a[1]
            );
            var rootKWs = XLSX.utils.json_to_sheet(this.ResultRootKeywords, {
              skipHeader: true,
            });
            let Heading = [["Keywords", "Count", "Percentage"]];
            XLSX.utils.sheet_add_aoa(rootKWs, Heading);
            XLSX.utils.sheet_add_json(rootKWs, this.ResultRootKeywords, { origin: 'A2', skipHeader: true });
            XLSX.utils.book_append_sheet(wb, rootKWs, "Root KWs");

            const MasterKeyords = XLSX.utils.json_to_sheet(this.MasterKwrds, {
              skipHeader: true,
            });
            
            XLSX.utils.book_append_sheet(wb, MasterKeyords, "Master Words");
            this.loading = false;
            XLSX.writeFile(wb, "Analyzed Keywords Sheet.xlsx");
          } catch (error) {
            this.popup = true;
            this.err = error;
            alert(error);
          }
        };

        myReader.readAsBinaryString(this.selectSheet);
      }
    },
    getSpilltedWords(arr) {
      var splitted = [];
      arr.forEach((element) => {
        element
          .filter((e) => e != "for")
          .forEach((word) => {
            splitted.push(word);
          });
      });
      return splitted;
    },
    dupCounts(arr) {
      var counts = {};
      arr.forEach(function (n) {
        // if property counts[n] doesn't exist, create it
        counts[n] = counts[n] || 0;
        // now increment it
        counts[n]++;
      });
      const objectToArray = (obj = {}) => {
        const res = [];
        const keys = Object.keys(obj);

        var total = this.getTotal(obj);

        for (var key of keys) {
          res.push([
            key,
            obj[key],
            ((obj[key] / total) * 100).toFixed(2) + "%",
          ]);
        }

        return res;
      };

      return objectToArray(counts);
    },
    getTotal(obj) {
      var total = 0;
      for (var property in obj) {
        total += obj[property];
      }
      return total;
    },
  },

  components: {
    // ClipLoader,
  },
};
</script>
<style>
.spinner {
  position: absolute;
  top: calc(40% - 25px);
  left: calc(50% - 25px);
}
h4 {
  font-family: "Roboto", sans-serif;
}
@import url("https://fonts.googleapis.com/css2?family=Roboto&display=swap");
</style>
