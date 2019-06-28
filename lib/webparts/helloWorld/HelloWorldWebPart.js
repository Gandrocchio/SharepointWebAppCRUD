"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
//import * as pnp from '../../../node_modules/sp-pnp-js';
var sp_1 = require("@pnp/sp");
require("@pnp/polyfill-ie11");
require("es6-object-assign/auto");
var HelloWorldWebPart_module_scss_1 = require("./HelloWorldWebPart.module.scss");
var strings = require("HelloWorldWebPartStrings");
var CRUDHelloWorld = /** @class */ (function (_super) {
    __extends(CRUDHelloWorld, _super);
    function CRUDHelloWorld() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CRUDHelloWorld.prototype.render = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () { return __awaiter(_this, void 0, void 0, function () {
                    var dialog;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                this.domElement.innerHTML = "\n      <div class=\"" + HelloWorldWebPart_module_scss_1.default.helloWorld + "\">\n        <div class=\"" + HelloWorldWebPart_module_scss_1.default.container + "\">\n          <div class=\"" + HelloWorldWebPart_module_scss_1.default.row + "\">\n            <div class=\"" + HelloWorldWebPart_module_scss_1.default.column + "\">\n              <span class=\"" + HelloWorldWebPart_module_scss_1.default.title + "\">Welcome to SharePoint!</span>\n              <p class=\"" + HelloWorldWebPart_module_scss_1.default.subTitle + "\">Customize SharePoint experiences using Web Parts.</p>\n              <p class=\"" + HelloWorldWebPart_module_scss_1.default.description + "\">" + sp_lodash_subset_1.escape(this.properties.description) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + HelloWorldWebPart_module_scss_1.default.button + "\">\n                <span class=\"" + HelloWorldWebPart_module_scss_1.default.label + "\">Learn more</span>\n              </a>\n            </div>\n            <div>\n                  <input id=\"Title\" placeholder=\"Title\" />          \n                  <input id=\"Nome\" placeholder=\"Nome\" />\n                  <input id=\"Cognome\" placeholder=\"Cognome\" />\n                  <button id=\"AddSPItem\" type=\"submit\">Aggiungi Elemento</button>\n                  <button id=\"UpdateSPItem\" type=\"submit\">Aggiorna Elemento</button>\n                  <button id=\"DeleteSPItem\" type=\"submit\">Cancella Elemento</button>\n            </div>\n            <br>\n            <div id =\"DivGetItems\" />\n          </div>\n        </div>\n       </div>\n      ";
                                this.AddEventListeners();
                                console.log("Inizio chiamata getSPItems");
                                dialog = SP.UI.ModalDialog.showWaitScreenWithNoClose('Processing...', 'questa modal scomparirà tra qualche secondo...', 130, 350);
                                console.log("Inizio render --> main()");
                                return [4 /*yield*/, this.getSPItemsAsync()];
                            case 1:
                                _a.sent();
                                console.log("Fine render --> main()");
                                dialog.close(SP.UI.DialogResult.OK);
                                return [2 /*return*/];
                        }
                    });
                }); });
                return [2 /*return*/];
            });
        });
    };
    Object.defineProperty(CRUDHelloWorld.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse("1.0");
        },
        enumerable: true,
        configurable: true
    });
    CRUDHelloWorld.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    CRUDHelloWorld.prototype.AddEventListeners = function () {
        var _this = this;
        document.getElementById("AddSPItem").addEventListener("click", function () { return _this.AddSPItem(); });
        document.getElementById("DeleteSPItem").addEventListener("click", function () { return _this.deleteSPItems(); });
        document.getElementById("UpdateSPItem").addEventListener("click", function () { return _this.UpdateSPItems(); });
    };
    // REST
    // Insert
    CRUDHelloWorld.prototype.AddSPItem = function () {
        sp_1.sp.web.lists.getByTitle("prova").items.add({
            Nome: document.getElementById("Nome")["value"],
            Cognome: document.getElementById("Cognome")["value"],
            Title: document.getElementById("Title")["value"]
        }).then(function (result) {
            alert("Operazone Completata");
            console.log(result.item);
        }).catch(function (errors) { return console.log(errors); });
    };
    // GET
    CRUDHelloWorld.prototype.getSPItems = function () {
        var _this = this;
        sp_1.sp.web.lists.getByTitle("prova").items.getAll().then(function (AllItems) {
            var stringHtml = "<div>";
            //let array: ISPList[];
            //array = [{ Nome: "Gabriele", Cognome: "Ascione", Title: "Prova" },];
            if (AllItems.length > 0) {
                AllItems.forEach(function (item) {
                    stringHtml += item.ID + " " + item.Cognome + "</br>";
                    console.log(item.ID + " " + item.Cognome);
                });
            }
            else {
                stringHtml += "Non ci sono elementi</br>";
            }
            stringHtml += "</div>";
            var listContainer = _this.domElement.querySelector('#DivGetItems');
            listContainer.innerHTML = stringHtml;
        }).catch(function (error) { return console.log(error); });
    };
    // GET Asincrona
    CRUDHelloWorld.prototype.getSPItemsAsync = function () {
        return __awaiter(this, void 0, void 0, function () {
            var AllItems, stringHtml, listContainer;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, sp_1.sp.web.lists.getByTitle("prova").items.getAll()];
                    case 1:
                        AllItems = _a.sent();
                        stringHtml = "<div>";
                        if (AllItems.length > 0) {
                            AllItems.forEach(function (item) {
                                stringHtml += item.ID + " " + item.Cognome + "</br>";
                                console.log(item.ID + " " + item.Cognome);
                            });
                        }
                        else {
                            stringHtml += "Non ci sono elementi</br>";
                        }
                        stringHtml += "</div>";
                        listContainer = this.domElement.querySelector('#DivGetItems');
                        listContainer.innerHTML = stringHtml;
                        return [2 /*return*/];
                }
            });
        });
    };
    // Delete
    CRUDHelloWorld.prototype.deleteSPItems = function () {
        var id = 1; // Ricerca tramite nome e cognome
        var list = sp_1.sp.web.lists.getByTitle("prova");
        list.items.getById(1).delete().then(function (_) { });
    };
    // Update
    CRUDHelloWorld.prototype.UpdateSPItems = function () {
        var id = 1;
        var list = sp_1.sp.web.lists.getByTitle("prova");
        list.items.getById(1).update({
            Nome: document.getElementById("Nome")["value"],
            Cognome: document.getElementById("Cognome")["value"],
            Title: document.getElementById("Title")["value"]
        }).then(function (i) {
            console.log(i);
        });
    };
    //Async Call e Promise
    CRUDHelloWorld.prototype.main = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log("Inizio main()");
                        return [4 /*yield*/, Promise.all([this.logConsole("1", 10000), this.logConsole("2", 10)])];
                    case 1:
                        _a.sent();
                        console.log("Fine main()");
                        return [2 /*return*/];
                }
            });
        });
    };
    CRUDHelloWorld.prototype.logConsole = function (s, timeout) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log("Inizio LogConsole");
                        return [4 /*yield*/, Promise.resolve(setTimeout(function () { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    console.log(s);
                                    console.log("Fine LogConsole");
                                    return [2 /*return*/];
                                });
                            }); }, timeout))];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    return CRUDHelloWorld;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = CRUDHelloWorld;

//# sourceMappingURL=HelloWorldWebPart.js.map
