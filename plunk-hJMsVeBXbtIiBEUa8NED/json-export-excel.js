(function() {
  'use strict';

  angular.module('ngJsonExportExcel', [])
    .directive('ngJsonExportExcel', function() {
      return {
        restrict: 'AE',
        scope: {
          data: '=',
          filename: '=?',
          reportFields: '=',
          nestedReportFields: '=',
          nestedDataProperty: "@"
        },
        link: function(scope, element) {
          scope.filename = !!scope.filename ? scope.filename : "export-excel";

          function generateFieldsAndHeaders(fieldsObject, fields, header) {
            angular.forEach(fieldsObject, function(field, key) {
              if (!field || !key) {
                throw new Error("error json report fields");
              }
              fields.push(key);
              header.push(field);
            });
            return {
              fields: fields,
              header: header
            };
          }
          var fieldsAndHeader = generateFieldsAndHeaders(scope.reportFields, [], []);
          var fields = fieldsAndHeader.fields,
            header = fieldsAndHeader.header;
          var nestedFieldsAndHeader = generateFieldsAndHeaders(scope.nestedReportFields, [], [""]);
          var nestedFields = nestedFieldsAndHeader.fields,
            nestedHeader = nestedFieldsAndHeader.header;

          function _convertToExcel(body, header) {
            return header + "\n" + body;
          }

          function _objectToString(object) {
            var output = "";
            angular.forEach(object, function(value, key) {
              output += key + ":" + value + " ";
            });

            return "'" + output + "'";
          }

          function generateFieldValues(list, rowItems, dataItem) {
            angular.forEach(list, function(field) {
              var data = "",
                fieldValue = "",
                curItem = null;
              if (field.indexOf(".")) {
                field = field.split(".");
                curItem = dataItem;
                // deep access to obect property
                angular.forEach(field, function(prop) {
                  if (curItem !== null && curItem !== undefined) {
                    curItem = curItem[prop];
                  }
                });
                data = curItem;
              } else {
                data = dataItem[field];
              }
              fieldValue = data !== null ? data : " ";
              if (fieldValue !== undefined && angular.isObject(fieldValue)) {
                fieldValue = _objectToString(fieldValue);
              }
              rowItems.push(fieldValue);
            });
            return rowItems;
          }

          function _bodyData() {
            var body = "";

            angular.forEach(scope.data, function(dataItem) {
              var rowItems = [];
              var nestedBody = "";
              rowItems = generateFieldValues(fields, rowItems, dataItem);
              //Nested Json body generation start 
              if (scope.nestedDataProperty && dataItem[scope.nestedDataProperty].length) {
                angular.forEach(dataItem[scope.nestedDataProperty], function(nestedDataItem) {
                  var nestedRowItems = [""];
                  nestedRowItems = generateFieldValues(nestedFields, nestedRowItems, nestedDataItem);
                  nestedBody += nestedRowItems.toString() + "\n";
                });
                var strData = _convertToExcel(nestedBody, nestedHeader);
                body += rowItems.toString() + "\n" + strData;
                ////Nested Json body generation end 
              } else {
                body += rowItems.toString() + "\n";
              }
            });
            return body;
          }

         // $timeout(function() {
            element.bind("click", function() {
              var bodyData = _bodyData();
              var strData = _convertToExcel(bodyData, header);
              var blob = new Blob([strData], {
                type: "text/plain;charset=utf-8"
              });

              return saveAs(blob, [scope.filename + ".csv"]);
            });
         // }, 1000);

        }
      };
    });
})();