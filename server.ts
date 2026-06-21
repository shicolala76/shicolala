/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import fs from "fs";
import AdmZip from "adm-zip";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";
import * as XLSX from "xlsx";
import {
  getAllSettings,
  getSetting,
  updateSetting,
  updateSettings,
  toggleSetting,
  resetAllSettings,
  exportSettings,
  importSettings
} from "./src/lib/settings";

dotenv.config({ override: false });

const apiKey = process.env.GEMINI_API_KEY;
const ai = new GoogleGenAI({
  apiKey: apiKey || "",
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Set high limits for base64 file payloads
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // Unified API endpoint to analyze any document & classify it automatically
  app.post("/api/invoices/analyze", async (req, res) => {
    try {
      const { fileBase64, fileName } = req.body;
      if (!fileBase64) {
        return res.status(400).json({ error: "لم يتم تقديم ملف للتحليل." });
      }

      if (!apiKey) {
        return res.status(500).json({ 
          error: "مفتاح API الخاص بـ Gemini غير مهيأ بالخادم. يرجى تفعيله في لوحة الأسرار (Secrets)." 
        });
      }

      // Handle Mime types and Base64 format extraction
      let mimeType = "image/png";
      let base64Data = fileBase64;

      const matches = fileBase64.match(/^data:([^;]+);base64,(.*)$/);
      if (matches && matches.length === 3) {
        mimeType = matches[1];
        base64Data = matches[2];
      } else {
        if (fileName) {
          const ext = fileName.split(".").pop()?.toLowerCase();
          if (ext === "pdf") {
            mimeType = "application/pdf";
          } else if (ext === "jpg" || ext === "jpeg") {
            mimeType = "image/jpeg";
          } else if (ext === "webp") {
            mimeType = "image/webp";
          }
        }
      }

      let isExcel = false;
      let excelText = "";
      const ext = fileName ? fileName.split(".").pop()?.toLowerCase() : "";
      if (ext === "xlsx" || ext === "xls" || ext === "csv" || mimeType.includes("sheet") || mimeType.includes("csv") || mimeType.includes("excel")) {
        isExcel = true;
      }

      if (isExcel) {
        try {
          const buffer = Buffer.from(base64Data, "base64");
          const workbook = XLSX.read(buffer, { type: "buffer" });
          let sheetsText: string[] = [];
          for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const csv = XLSX.utils.sheet_to_csv(sheet);
            sheetsText.push(`ورقة العمل [Sheet: ${sheetName}]:\n${csv}`);
          }
          excelText = sheetsText.join("\n\n");
        } catch (excelErr) {
          console.error("Server-side Excel parse error:", excelErr);
          isExcel = false;
        }
      }

      const contents: any[] = [];
      if (isExcel) {
        contents.push({
          text: `هذا هو محتوى جدول الإكسل المرفوع بالكامل لساحة ومطحنة الجلخ:\n\n${excelText}`
        });
      } else {
        contents.push({
          inlineData: {
            mimeType: mimeType,
            data: base64Data,
          },
        });
      }

      contents.push(
        `قم بتحليل وتدقيق هذه الفاتورة أو الإيصال أو شهادة الوزن، وصنفها بدقة إما كمسند شراء (purchase) أو بيع (sale) أو مصروف تشغيلي (expense) أو معاملة بنكية وحسابية (bank).
ملاحظات هامة جداً لضبط البيانات بذكاء:
1. اسم الملف المرفوع هو: "${fileName || ''}". إذا كان اسم الملف يحتوي على اسم شركة أو عميل أو مورد (مثل "عز"، "بشاي"، "أبو حماد"، إلخ)، فيرجى اعتباره اسم العميل أو المورد (partyName) بدلاً من استخراجه بشكل خاطئ من تفاصيل ورقية أخرى.
2. استخرج اسم المقاول (contractorName) إذا ورد في المستند أو اسم الملف.
3. استخرج اسم السائق (driverName) ورقم السيارة أو الشاحنة (carNumber) إذا وردت في المستند أو اسم الملف.
استخرج كافة قيم ومبالغ وتفاصيل المعاملة بدقة بالغة محاسبياً.`
      );

      // Call Gemini 3.5 Flash to classify and extract
      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: contents,
        config: {
          systemInstruction: `أنت كبير المحاسبين والمدققين الماليين في شركة الهضبة لتجارة الجلخ والمعادن بمصر.
مهمتك المحورية هي استلام إيصالات ومعاملات وفواتير وجداول الساحة (صور أو ملفات PDF أو جداول إكسيل) وتصنيفها كالتالي:
1. "purchase" (فاتورة شراء جلخ واردة): عندما تقوم الشركة بشراء جلخ من تاجر أو مورد.
2. "sale" (فاتورة بيع جلخ صادرة): عندما نبيع الجلخ لمصانع حديد عز أو بشاي للصلب أو المصرية للمعادن.
3. "expense" (مصروف تشغيلي): مثل فواتير النولون والسيارات، يوميات الورشة والعمال، فواتير الكهرباء، الصيانة، الخ.
4. "bank" (معاملة بنكية أو إيداع وسحب): حركة دفع أو تحويل مالي أو إيداع واردة في حسابات البنك أو كشوفات إكسيل.

قواعد الاستخراج الدقيقة:
- قارن اسم العميل أو المورد الموجود مع قاعدة أسماء المستودع بمصر.
- طابق مادة الجلخ دائماً مع المعرف الفريد التالي بدقة، وحدد معرف الصنف المناسب ('itemId'):
  * 'it-1': جلخ حديد تسليح مميز (Premium Iron Slag)
- إذا كانت الفاتورة بالأوزان والأسعار للجلخ، حوِّل أي وزن بالطن إلى الكيلوجرام بضربه في 1000 (مثال: 2.5 طن تساوي 2500 كجم) واكتب السعر بالنسبة للطن (مثال: سعر الكيلو 2 جنيه يعني سعر الطن 2000 جنيه).
- أرجع النتيجة في الهيكل المناسب تماماً.`,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              transactionType: {
                type: Type.STRING,
                description: "النوع الحرفي للعمليات: 'purchase' أو 'sale' أو 'expense' أو 'bank'"
              },
              confidenceReasoning: {
                type: Type.STRING,
                description: "بين يدي الفحص: شرح مفصل باللغة العربية للبيانات المستخرجة وسبب التصنيف"
              },
              expense: {
                type: Type.OBJECT,
                description: "تفاصيل المصروف وتملأ فقط إذا كان النوع expense",
                properties: {
                  amount: { type: Type.NUMBER, description: "إجمالي المبلغ كقيمة رقمية صافية" },
                  categoryKey: { type: Type.STRING, description: "أحد الكلمات: Transportation, Labor, Rent, Utilities, Commissions, Maintenance, Taxes, Other" },
                  subCategory: { type: Type.STRING, description: "التصنيف الفرعي المحدد بالعربية" },
                  description: { type: Type.STRING, description: "شرح المصروف والغرض منه باللغة العربية" },
                  date: { type: Type.STRING, description: "التاريخ بالصيغة القياسية YYYY-MM-DD" },
                  paymentMethod: { type: Type.STRING, description: "طريقة السداد: cash أو bank أو cheque" },
                  receiptNumber: { type: Type.STRING, description: "رقم الفاتورة أو الدفتر" },
                  supplierName: { type: Type.STRING, description: "اسم الجهة المستلمة للمال" },
                  notes: { type: Type.STRING, description: "أي ملاحظات إضافية" }
                }
              },
              invoice: {
                type: Type.OBJECT,
                description: "تفاصيل الفاتورة وتملأ إذا كان نوع المعاملة purchase أو sale",
                properties: {
                  partyName: { type: Type.STRING, description: "اسم الشريك التجاري المستلم أو المورد" },
                  contractorName: { type: Type.STRING, description: "اسم المقاول أو جهة النقل والتحميل الفوقية" },
                  driverName: { type: Type.STRING, description: "اسم سائق الشاحنة والتريلا إن وجد" },
                  carNumber: { type: Type.STRING, description: "رقم لوحة أو ترخيص السيارة أو سيارة النقل" },
                  date: { type: Type.STRING, description: "تاريخ الفاتورة بالصيغة YYYY-MM-DD" },
                  paymentType: { type: Type.STRING, description: "نوع السداد: cash أو credit" },
                  paidAmount: { type: Type.NUMBER, description: "المبلغ الذي تم سداده نقداً بالفعل" },
                  notes: { type: Type.STRING, description: "ملاحظات الفاتورة إن وجدت" },
                  details: {
                    type: Type.ARRAY,
                    description: "تفاصيل وتفصيلات المواد والحديد والنحاس الواردة في المستند",
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        itemId: { type: Type.STRING, description: "معرف الصنف في المستودع: 'it-1', 'it-2', 'it-3', 'it-4', or 'it-5'" },
                        itemName: { type: Type.STRING, description: "اسم المادة بالعربية" },
                        weightKg: { type: Type.NUMBER, description: "الوزن المستخرج بالكيلوجرام (طابق بدقة!)" },
                        pricePerTon: { type: Type.NUMBER, description: "السعر المستهدف للطن الواحد (1000 كجم)" }
                      },
                      required: ["itemId", "itemName", "weightKg", "pricePerTon"]
                    }
                  }
                }
              },
              bankTransaction: {
                type: Type.OBJECT,
                description: "تفاصيل المعاملة البنكية وتملأ فقط إذا كان النوع bank",
                properties: {
                  bankName: { type: Type.STRING, description: "اسم البنك المفترض أو المذكور في التحويل مثل البنك الأهلي، بنك مصر، فودافون كاش، CIB الخ" },
                  type: { type: Type.STRING, description: "نوع المعاملة: deposit (إيداع/وارد) أو withdrawal (سحب/صادر)" },
                  amount: { type: Type.NUMBER, description: "مبلغ التحويل كقيمة رقمية" },
                  date: { type: Type.STRING, description: "التاريخ بالصيغة القياسية YYYY-MM-DD" },
                  description: { type: Type.STRING, description: "بيان التحويل أو الغرض منه باللغة العربية" },
                  notes: { type: Type.STRING, description: "أية ملاحظات أو رقم مرجع التحويل" }
                }
              }
            },
            required: ["transactionType", "confidenceReasoning"]
          }
        }
      });

      const textResult = response.text || "{}";
      try {
        const parsedData = JSON.parse(textResult.trim());
        res.json(parsedData);
      } catch {
        res.status(500).json({ 
          error: "فشلت عملية تهيئة وقراءة البيانات من مخرجات النموذج الذكي.", 
          raw: textResult 
        });
      }
    } catch (err: any) {
      console.error("AI Unified analyze error:", err);
      res.status(500).json({ error: err.message || "خطأ داخلي أثناء معالجة وقراءة الفاتورة." });
    }
  });

  // API endpoint to analyze invoices & receipts using AI (image or PDF)
  app.post("/api/expenses/analyze", async (req, res) => {
    try {
      const { fileBase64, fileName } = req.body;
      if (!fileBase64) {
        return res.status(400).json({ error: "لم يتم تقديم ملف للتحليل." });
      }

      if (!apiKey) {
        return res.status(500).json({ 
          error: "مفتاح API الخاص بـ Gemini غير مهيأ بالخادم. يرجى تفعيله في لوحة الأسرار (Secrets)." 
        });
      }

      // Handle Mime types and Base64 format extraction
      let mimeType = "image/png";
      let base64Data = fileBase64;

      const matches = fileBase64.match(/^data:([^;]+);base64,(.*)$/);
      if (matches && matches.length === 3) {
        mimeType = matches[1];
        base64Data = matches[2];
      } else {
        if (fileName) {
          const ext = fileName.split(".").pop()?.toLowerCase();
          if (ext === "pdf") {
            mimeType = "application/pdf";
          } else if (ext === "jpg" || ext === "jpeg") {
            mimeType = "image/jpeg";
          } else if (ext === "webp") {
            mimeType = "image/webp";
          }
        }
      }

      // Call Gemini with the invoice file and structured schemas
      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: [
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Data,
            },
          },
          "قم بقراءة وتحليل هذه الفاتورة أو المستند بذكاء واستخرج جميع البيانات المالية والتنظيمية باللغة العربية كمسؤول حسابات لمستودع خردة ومعادن.",
        ],
        config: {
          systemInstruction: `أنت خبير حسابات وتدقيق في شركة الهضبة لتجارة الخردة والمعادن بمصر بمستوى احترافي عالي.
مهمتك هي قراءة إيصالات ومستندات وفواتير المصاريف المرفوعة (صور أو ملفات PDF) واستخراج قيم الحقول باللغة العربية بصيغة منظمة تماماً.

يجب مطابقة حقل التصنيف الرئيس 'categoryKey' بدقة تامة مع أحد التصنيفات الثمانية المدعومة بالنظام فقط:
- 'Transportation' (نقل وشحن)
- 'Labor' (عمالة)
- 'Rent' (إيجارات)
- 'Utilities' (مرافق)
- 'Commissions' (عمولات)
- 'Maintenance' (صيانة)
- 'Taxes' (ضرائب ورسوم)
- 'Other' (أخرى)

واختر تصنيفاً فرعياً مناسباً 'subCategory' باللغة العربية من الأمثلة الشائعة أو القريبة للمستند:
- Transportation: 'إيجار سيارات نقل'، 'وقود (سولار/بنزين)'، 'صيانة سيارات'، 'تحميل وتفريغ'
- Labor: 'رواتب شهرية'، 'مكافآت وحوافز'، 'أجر يومي (عمال موسميين)'
- Utilities: 'كهرباء'، 'مياه'، 'غاز'، 'إنترنت وتليفونات'
- Other: 'ضيافة'، 'قرطاسية'، 'سفر وانتقالات'، إلخ.

ضع المبلغ المستخرج في الحقل 'amount' كقيمة رقمية صافية.
استخرج التاريخ بالصيغة القياسية 'YYYY-MM-DD' في حقل 'date'.
اكتب وصفاً معبراً ودقيقاً للفاتورة في 'description'، واستخرج اسم المورد 'supplierName'، ورقم الإيصال 'receiptNumber'، وأية ملاحظات أخرى إضافية في حقل 'notes'.`,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              amount: {
                type: Type.NUMBER,
                description: "إجمالي المبلغ المستخرج كرقم"
              },
              categoryKey: {
                type: Type.STRING,
                description: "مفتاح التصنيف ويجب أن يكون حرفياً واحد من: 'Transportation', 'Labor', 'Rent', 'Utilities', 'Commissions', 'Maintenance', 'Taxes', 'Other'"
              },
              subCategory: {
                type: Type.STRING,
                description: "التصنيف الفرعي للمصروف باللغة العربية"
              },
              description: {
                type: Type.STRING,
                description: "وصف واضح ومختصر للمصروف باللغة العربية"
              },
              date: {
                type: Type.STRING,
                description: "التاريخ بالصيغة YYYY-MM-DD"
              },
              receiptNumber: {
                type: Type.STRING,
                description: "رقم الفاتورة أو الإيصال أو الدفتر إن وجد"
              },
              supplierName: {
                type: Type.STRING,
                description: "اسم الجهة المستفيدة أو المورد أو المحل"
              },
              notes: {
                type: Type.STRING,
                description: "أي ملاحظات محاسبية هامة إضافية"
              }
            },
            required: ["amount", "categoryKey", "subCategory", "description", "date"]
          }
        }
      });

      const textResult = response.text || "{}";
      try {
        const parsedData = JSON.parse(textResult.trim());
        res.json(parsedData);
      } catch {
        res.status(500).json({ 
          error: "فشلت عملية تهيئة وقراءة البيانات من مخرجات النموذج الذكي.", 
          raw: textResult 
        });
      }
    } catch (err: any) {
      console.error("AI analyze error:", err);
      res.status(500).json({ error: err.message || "خطأ داخلي أثناء معالجة وقراءة الفاتورة." });
    }
  });

  // ============= Settings API =============

  // GET /api/settings - جلب جميع الإعدادات
  app.get('/api/settings', (req, res) => {
    try {
      const settings = getAllSettings();
      // Keep support for both 'data' and 'settings' keys for total backward-compatibility
      res.json({ success: true, data: settings, settings });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // GET /api/settings/export - تصدير الإعدادات كـ JSON
  app.get('/api/settings/export', (req, res) => {
    try {
      const jsonData = exportSettings();
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Content-Disposition', 'attachment; filename="settings_backup.json"');
      res.send(jsonData);
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // GET /api/settings/:key - جلب إعداد معين
  app.get('/api/settings/:key', (req, res) => {
    try {
      const { key } = req.params;
      const value = getSetting(key);
      res.json({ success: true, data: { key, value } });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // POST /api/settings/batch - تحديث عدة إعدادات دفعة واحدة
  app.post('/api/settings/batch', (req, res) => {
    try {
      const { settings, updatedBy } = req.body;
      const success = updateSettings(settings, updatedBy);
      res.json({ success });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // POST /api/settings/reset - إعادة تعيين الإعدادات
  app.post('/api/settings/reset', (req, res) => {
    try {
      const success = resetAllSettings();
      res.json({ success });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // POST /api/settings/import - استيراد الإعدادات
  app.post('/api/settings/import', (req, res) => {
    try {
      const { data, updatedBy } = req.body;
      const success = importSettings(data, updatedBy);
      res.json({ success });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // POST /api/settings/toggle - تشغيل/إلغاء مفتاح
  app.post("/api/settings/toggle", (req, res) => {
    try {
      const { key, updatedBy } = req.body;
      if (!key) {
        return res.status(400).json({ success: false, error: "المفتاح (key) مطلوب." });
      }
      const success = toggleSetting(key, updatedBy);
      res.json({ success });
    } catch (err: any) {
      res.status(500).json({ success: false, error: err.message });
    }
  });

  // PUT /api/settings - للتوافق مع الواجهات السابقة
  app.put("/api/settings", (req, res) => {
    try {
      const { key, value, settings, updatedBy } = req.body;
      if (Array.isArray(settings)) {
        const success = updateSettings(settings, updatedBy);
        return res.json({ success });
      } else if (key && value !== undefined) {
        const success = updateSetting(key, String(value), updatedBy);
        return res.json({ success });
      }
      return res.status(400).json({ success: false, error: "بيانات الإدخال غير صالحة." });
    } catch (err: any) {
      res.status(500).json({ success: false, error: err.message });
    }
  });

  // POST /api/settings/:key - تحديث إعداد معين
  app.post('/api/settings/:key', (req, res) => {
    try {
      const { key } = req.params;
      const { value, updatedBy } = req.body;
      const success = updateSetting(key, value, updatedBy);
      res.json({ success, data: { key, value } });
    } catch (error: any) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // API endpoint to compress and download the workspace as a ZIP file
  app.get("/api/download-zip", (req, res) => {
    try {
      const zip = new AdmZip();
      const rootPath = process.cwd();

      function addFiles(currentDir: string) {
        const files = fs.readdirSync(currentDir);
        for (const file of files) {
          const fullPath = path.join(currentDir, file);
          const relPath = path.relative(rootPath, fullPath);

          // Skip large or unnecessary directories & actual hidden secrets
          if (
            file === "node_modules" ||
            file === "dist" ||
            file === ".git" ||
            file === ".cache" ||
            file === ".env" ||
            file === "package-lock.json" ||
            file === ".DS_Store"
          ) {
            continue;
          }

          const stat = fs.statSync(fullPath);
          if (stat.isDirectory()) {
            addFiles(fullPath);
          } else {
            const zipPath = path.dirname(relPath);
            const targetFolderInZip = zipPath === "." ? "" : zipPath;
            zip.addLocalFile(fullPath, targetFolderInZip);
          }
        }
      }

      addFiles(rootPath);

      const zipBuffer = zip.toBuffer();
      res.setHeader("Content-Type", "application/zip");
      res.setHeader("Content-Disposition", 'attachment; filename="elhadaba-scrap-metal.zip"');
      res.send(zipBuffer);
    } catch (err: any) {
      console.error("Zip compression error:", err);
      res.status(500).json({ error: "فشلت عملية تجميع وضغط ملفات المشروع: " + err.message });
    }
  });

  // Configure Vite or Serve SPA static files
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Express custom server running on port ${PORT}`);
  });
}

startServer();
