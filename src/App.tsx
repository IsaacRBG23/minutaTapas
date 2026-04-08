/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle2, X, AlertCircle, FileDown, Calendar } from 'lucide-react';
import * as xlsx from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

export default function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [dataPreview, setDataPreview] = useState<any[]>([]);
  const [fullData, setFullData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [workbook, setWorkbook] = useState<xlsx.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [folio, setFolio] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    setSelectedDate(`${yyyy}-${mm}-${dd}`);
  }, []);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      processFile(droppedFile);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  const processFile = (selectedFile: File) => {
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'text/csv'
    ];
    
    const fileExtension = selectedFile.name.split('.').pop()?.toLowerCase();
    const isValidExtension = fileExtension === 'xlsx' || fileExtension === 'xls' || fileExtension === 'csv';

    if (!validTypes.includes(selectedFile.type) && !isValidExtension) {
      setError('Por favor, sube un archivo Excel válido (.xlsx, .xls o .csv)');
      return;
    }

    setFile(selectedFile);
    readExcel(selectedFile);
  };

  const readExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const wb = xlsx.read(data, { type: 'binary' });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);
        
        if (wb.SheetNames.length > 0) {
          const firstSheet = wb.SheetNames[0];
          setSelectedSheet(firstSheet);
          loadSheetData(wb, firstSheet);
        }
      } catch (err) {
        setError('Hubo un error al leer el archivo. Asegúrate de que no esté corrupto.');
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const loadSheetData = (wb: xlsx.WorkBook, sheetName: string) => {
    const worksheet = wb.Sheets[sheetName];
    
    // Extraer folio de la celda I3
    const cellI3 = worksheet['I3'];
    if (cellI3 && (cellI3.v !== undefined || cellI3.w !== undefined)) {
      setFolio(String(cellI3.v !== undefined ? cellI3.v : cellI3.w).trim());
    } else {
      setFolio('S/N');
    }

    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as any[][];
    
    // La tabla comienza en B6 (fila 6, índice 5) hasta I6 (columna B es índice 1, I es índice 8)
    if (jsonData.length > 5) {
      const headerRow = jsonData[5] || [];
      const headers: string[] = [];
      
      // Extraer encabezados de B a I (índices 1 a 8)
      for (let i = 1; i <= 8; i++) {
        headers.push(headerRow[i] !== undefined && headerRow[i] !== '' ? String(headerRow[i]) : `Col ${i}`);
      }
      
      // Los datos comienzan en B7 (fila 7, índice 6)
      const allRows = jsonData.slice(6)
        .map(row => {
          const extracted = [];
          for (let i = 1; i <= 8; i++) {
            extracted.push(row[i] !== undefined ? row[i] : '');
          }
          return extracted;
        })
        .filter(row => {
          // Filtrar filas donde las columnas C a I (índices 1 a 7 de extracted) estén completamente vacías
          const dataCols = row.slice(1, 8);
          return dataCols.some(cell => String(cell).trim() !== '');
        });
      
      const previewRows = allRows.slice(0, 5); // Preview first 5 rows
      
      setColumns(headers);
      setDataPreview(previewRows);
      setFullData(allRows);
    } else {
      setColumns([]);
      setDataPreview([]);
      setFullData([]);
    }
  };

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newSheet = e.target.value;
    setSelectedSheet(newSheet);
    if (workbook) {
      loadSheetData(workbook, newSheet);
    }
  };

  const formatExcelDate = (val: any) => {
    if (val === undefined || val === null || val === '') return '';
    
    const numVal = Number(val);
    if (!isNaN(numVal) && numVal > 10000) {
      const date = new Date(Math.round((numVal - 25569) * 86400 * 1000));
      const d = date.getUTCDate().toString().padStart(2, '0');
      const m = (date.getUTCMonth() + 1).toString().padStart(2, '0');
      return `${d}-${m}`;
    }
    
    const strVal = String(val).trim();
    const parts = strVal.split(/[-/]/);
    if (parts.length >= 2) {
      if (strVal.includes('-') && parts[0].length === 4) {
        return `${parts[2].substring(0, 2)}-${parts[1]}`;
      }
      return `${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
    }
    return strVal;
  };

  const formatKm = (val: any) => {
    if (val === undefined || val === null || String(val).trim() === '') return '';
    
    const strVal = String(val).replace(/,/g, '');
    const numVal = Number(strVal);
    
    if (isNaN(numVal)) return String(val);

    const isNegative = numVal < 0;
    const absVal = Math.abs(numVal);
    
    const thousands = Math.floor(absVal / 1000);
    const remainder = absVal % 1000;
    
    const thousandsStr = String(thousands).padStart(2, '0');
    
    const remainderParts = remainder.toFixed(2).split('.');
    const remainderInt = remainderParts[0].padStart(3, '0');
    const remainderDec = remainderParts[1];
    
    const sign = isNegative ? '-' : '';
    
    return `${sign}${thousandsStr}+${remainderInt}.${remainderDec}`;
  };

  const generatePDF = () => {
    const doc = new jsPDF();
    const tableColumn = ['N°', 'Fecha', 'Cod', 'Vialidad', 'Sección', 'Km', 'Carril', 'Falla'];
    
    const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    const dateObj = new Date(selectedDate + 'T12:00:00');
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = months[dateObj.getMonth()];
    const year = dateObj.getFullYear();
    
    const headerText = `Siendo el día ${day} de ${month} de ${year}, deribado de los recorridos de inspección de trabajos de reposicion y renivelacion de tapas se observaron los siguientes trabajos a realizar.`;

    // Dividir los datos en bloques de 25 registros
    const chunkSize = 25;
    const chunks = [];
    for (let i = 0; i < fullData.length; i += chunkSize) {
      chunks.push(fullData.slice(i, i + chunkSize));
    }
    
    // Asegurar que al menos haya una página si no hay datos
    if (chunks.length === 0) {
      chunks.push([]);
    }

    chunks.forEach((chunk, pageIndex) => {
      if (pageIndex > 0) {
        doc.addPage();
      }
      
      // Title
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.text('Minuta de Trabajo Obras de Drenaje (tapas)', 105, 20, { align: 'center' });
      
      // Folio / ID
      doc.setTextColor(200, 0, 0);
      doc.setFontSize(14);
      doc.text(folio, 195, 20, { align: 'right' });
      
      // Header Text
      doc.setTextColor(0, 0, 0);
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.text(headerText, 14, 30, { maxWidth: 180 });
      
      // Preparar exactamente 25 filas para la página actual
      const tableRows: any[] = [];
      for (let i = 0; i < 25; i++) {
        if (i < chunk.length) {
          const row = chunk[i];
          tableRows.push([
            String(i + 1), // N° siempre reinicia de 1 a 25
            formatExcelDate(row[1]), // Fecha (C)
            row[2] !== undefined ? String(row[2]) : '', // Cod (D)
            row[3] !== undefined ? String(row[3]) : '', // Vialidad (E)
            row[4] !== undefined ? String(row[4]) : '', // Sección (F)
            formatKm(row[5]), // Km (G)
            row[6] !== undefined ? String(row[6]) : '', // Carril (H)
            row[7] !== undefined ? String(row[7]) : ''  // Falla (I)
          ]);
        } else {
          // Rellenar con filas vacías hasta llegar a 25
          tableRows.push([String(i + 1), '', '', '', '', '', '', '']);
        }
      }

      autoTable(doc, {
        startY: 40,
        head: [
          [{ content: 'Relación de trabajos de reposicion y renivelacion de tapas a realizar:', colSpan: 8, styles: { halign: 'center', fillColor: [200, 200, 200], textColor: [0, 0, 0], fontStyle: 'bold' } }],
          tableColumn
        ],
        body: tableRows,
        theme: 'grid',
        headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center', fontSize: 8, lineWidth: 0.1, lineColor: [0, 0, 0] },
        bodyStyles: { fontSize: 8, minCellHeight: 6, textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0], halign: 'center' },
        columnStyles: {
          0: { cellWidth: 10, halign: 'center' },
          1: { cellWidth: 20, halign: 'center' },
          2: { cellWidth: 15, halign: 'center' },
          3: { cellWidth: 45, halign: 'center' },
          4: { cellWidth: 20, halign: 'center' },
          5: { cellWidth: 15, halign: 'center' },
          6: { cellWidth: 20, halign: 'center' },
          7: { cellWidth: 'auto', halign: 'center' }
        },
        margin: { top: 40, left: 14, right: 14 }
      });

      const finalY = (doc as any).lastAutoTable.finalY || 40;

      // Footer Text
      doc.setFontSize(10);
      const footerText = 'Siendo las ______horas del día____ de ______________ de ____ firman de conformidad en el recorrido de inspección las personas que se enlistan al calce de la presente minuta.';
      doc.text(footerText, 14, finalY + 10, { maxWidth: 180 });

      // Signatures Table
      autoTable(doc, {
        startY: finalY + 20,
        head: [['AXIS', 'CONTRATISTA']],
        body: [
          ['\n\n\nNombre:', '\n\n\nNombre:\nEmpresa:']
        ],
        theme: 'grid',
        headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center', fontSize: 9, lineWidth: 0.1, lineColor: [0, 0, 0] },
        bodyStyles: { fontSize: 9, minCellHeight: 25, valign: 'bottom', textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0] },
        columnStyles: {
          0: { cellWidth: 91 },
          1: { cellWidth: 91 }
        },
        margin: { left: 14, right: 14 }
      });
    });

    const safeFolio = folio ? folio.replace(/[\s\/\\]/g, '_') : 'S_N';
    doc.save(`Minuta_de_Trabajo_${safeFolio}.pdf`);
  };

  const resetState = () => {
    setFile(null);
    setWorkbook(null);
    setSheetNames([]);
    setSelectedSheet('');
    setFolio('');
    setDataPreview([]);
    setFullData([]);
    setColumns([]);
    setError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col items-center py-12 px-4 sm:px-6 lg:px-8 font-sans text-gray-900">
      <div className="w-full max-w-3xl space-y-8">
        
        {/* Header */}
        <div className="text-center">
          <h1 className="text-3xl font-semibold tracking-tight text-gray-900">
            Carga de Datos
          </h1>
          <p className="mt-2 text-sm text-gray-500">
            Sube tu archivo Excel para previsualizar y procesar la información.
          </p>
        </div>

        {/* Upload Zone */}
        {!file && (
          <div
            className={`relative flex flex-col items-center justify-center w-full p-12 mt-8 border-2 border-dashed rounded-2xl transition-all duration-200 ease-in-out cursor-pointer ${
              isDragging 
                ? 'border-blue-500 bg-blue-50' 
                : 'border-gray-300 bg-white hover:bg-gray-50 hover:border-gray-400'
            }`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".xlsx, .xls, .csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
              className="hidden"
            />
            
            <div className={`p-4 rounded-full mb-4 ${isDragging ? 'bg-blue-100 text-blue-600' : 'bg-gray-100 text-gray-500'}`}>
              <UploadCloud className="w-8 h-8" />
            </div>
            
            <h3 className="text-lg font-medium text-gray-900 mb-1">
              Haz clic o arrastra tu archivo aquí
            </h3>
            <p className="text-sm text-gray-500 text-center max-w-xs">
              Soporta archivos .xlsx, .xls y .csv hasta 10MB
            </p>
          </div>
        )}

        {/* Error Message */}
        {error && (
          <div className="flex items-center p-4 mt-4 text-sm text-red-800 border border-red-200 rounded-xl bg-red-50">
            <AlertCircle className="w-5 h-5 mr-3 flex-shrink-0" />
            <span>{error}</span>
          </div>
        )}

        {/* Success & Preview State */}
        {file && !error && (
          <div className="mt-8 bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
            {/* File Info Header */}
            <div className="flex items-center justify-between p-6 border-b border-gray-100 bg-gray-50/50">
              <div className="flex items-center space-x-4">
                <div className="p-3 bg-green-100 text-green-600 rounded-xl">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="text-sm font-medium text-gray-900 flex items-center">
                    {file.name}
                    <CheckCircle2 className="w-4 h-4 text-green-500 ml-2" />
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    {(file.size / 1024 / 1024).toFixed(2)} MB
                  </p>
                </div>
              </div>
              <button
                onClick={resetState}
                className="p-2 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-lg transition-colors"
                title="Eliminar archivo"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Sheet Selector */}
            {sheetNames.length > 1 && (
              <div className="px-6 pt-6 pb-2 border-b border-gray-100">
                <label htmlFor="sheet-select" className="block text-sm font-medium text-gray-700 mb-2">
                  Selecciona la hoja de cálculo a procesar
                </label>
                <select
                  id="sheet-select"
                  value={selectedSheet}
                  onChange={handleSheetChange}
                  className="block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-gray-900 focus:border-gray-900 sm:text-sm rounded-xl border bg-white"
                >
                  {sheetNames.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Date Selector */}
            <div className="px-6 pt-4 pb-2 border-b border-gray-100">
              <label htmlFor="date-select" className="block text-sm font-medium text-gray-700 mb-2">
                Fecha para el reporte
              </label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <Calendar className="h-5 w-5 text-gray-400" />
                </div>
                <input
                  type="date"
                  id="date-select"
                  value={selectedDate}
                  onChange={(e) => setSelectedDate(e.target.value)}
                  className="block w-full pl-10 pr-3 py-2 border border-gray-300 rounded-xl focus:outline-none focus:ring-gray-900 focus:border-gray-900 sm:text-sm bg-white"
                />
              </div>
            </div>

            {/* Data Preview Table */}
            <div className="p-6">
              {columns.length > 0 ? (
                <>
                  <h4 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-4">
                    Vista previa de datos (primeras 5 filas)
                  </h4>
                <div className="overflow-x-auto rounded-lg border border-gray-200">
                  <table className="min-w-full divide-y divide-gray-200 text-sm">
                    <thead className="bg-gray-50">
                      <tr>
                        {columns.map((col, i) => (
                          <th
                            key={i}
                            scope="col"
                            className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider whitespace-nowrap"
                          >
                            {col || `Columna ${i + 1}`}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-gray-200">
                      {dataPreview.map((row, rowIndex) => (
                        <tr key={rowIndex} className="hover:bg-gray-50 transition-colors">
                          {columns.map((_, colIndex) => (
                            <td
                              key={colIndex}
                              className="px-6 py-4 whitespace-nowrap text-gray-600"
                            >
                              {row[colIndex] !== undefined && row[colIndex] !== '' ? String(row[colIndex]) : '-'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {dataPreview.length === 0 && (
                  <p className="text-sm text-gray-500 mt-4 text-center">
                    El archivo parece estar vacío.
                  </p>
                )}
                
                <div className="mt-6 flex justify-end">
                  <button 
                    onClick={generatePDF}
                    className="flex items-center px-6 py-2.5 bg-gray-900 text-white text-sm font-medium rounded-xl hover:bg-gray-800 transition-colors shadow-sm"
                  >
                    <FileDown className="w-4 h-4 mr-2" />
                    Generar PDF
                  </button>
                </div>
                </>
              ) : (
                <p className="text-sm text-gray-500 text-center py-4">
                  La hoja seleccionada está vacía o no tiene el formato esperado.
                </p>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
