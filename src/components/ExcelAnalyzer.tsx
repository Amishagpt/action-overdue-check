import React, { useState, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, Calendar } from 'lucide-react';
import { useToast } from '@/hooks/use-toast';
import * as XLSX from 'xlsx';

interface AnalysisResult {
  total_rows: number;
  assigned_count: number;
  assigned_pct: number;
  overdue_count: number;
  overdue_pct_of_assigned: number;
  today_iso: string;
  timezone: string;
  columns_used: { action: string; due_date: string };
  notes: string[];
  summary: string;
}

const ExcelAnalyzer = () => {
  const [dragActive, setDragActive] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { toast } = useToast();

  const isAssigned = (value: any): boolean => {
    if (value === null || value === undefined || value === '') return false;
    const str = String(value).toLowerCase().trim();
    return ['yes', 'true', 'assigned', 'done', '1'].includes(str) || 
           (str !== 'no' && str !== 'false' && str !== 'unassigned' && str !== '0' && str !== '');
  };

  const parseDate = (dateValue: any): Date | null => {
    if (!dateValue) return null;
    
    // Handle Excel serial dates
    if (typeof dateValue === 'number') {
      const excelEpoch = new Date(1900, 0, 1);
      const days = dateValue - 2; // Excel starts from 1900-01-01, but has leap year bug
      return new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000);
    }
    
    // Handle date strings
    const parsed = new Date(dateValue);
    return isNaN(parsed.getTime()) ? null : parsed;
  };

  const getTodayInKolkata = (): Date => {
    const now = new Date();
    const kolkataTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
    return new Date(kolkataTime.getFullYear(), kolkataTime.getMonth(), kolkataTime.getDate());
  };

  const findColumns = (worksheet: XLSX.WorkSheet): { actionCol: string | null; dueDateCol: string | null } => {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    let actionCol = null;
    let dueDateCol = null;

    // Check first row for headers
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      const cell = worksheet[cellAddress];
      if (!cell) continue;

      const header = String(cell.w || cell.v || '').toLowerCase().trim();
      if (header.includes('action') && !actionCol) {
        actionCol = XLSX.utils.encode_col(col);
      } else if ((header.includes('due') && header.includes('date')) || header === 'due date' && !dueDateCol) {
        dueDateCol = XLSX.utils.encode_col(col);
      }
    }

    return { actionCol, dueDateCol };
  };

  const analyzeExcel = useCallback(async (file: File) => {
    setIsAnalyzing(true);
    setError(null);
    setResult(null);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

      const { actionCol, dueDateCol } = findColumns(worksheet);
      
      if (!actionCol) {
        throw new Error('Action column not found. Please ensure your Excel file has an "Action" column.');
      }

      const notes: string[] = [];
      if (!dueDateCol) {
        notes.push('Due Date column not found. Overdue analysis skipped.');
      }

      const today = getTodayInKolkata();
      let totalRows = 0;
      let assignedCount = 0;
      let overdueCount = 0;

      // Analyze data (skip header row)
      for (let row = 1; row <= range.e.r; row++) {
        const actionCellAddr = `${actionCol}${row + 1}`;
        const actionCell = worksheet[actionCellAddr];
        
        if (!actionCell && !dueDateCol) continue; // Skip completely empty rows
        
        totalRows++;
        const isActionAssigned = isAssigned(actionCell?.w || actionCell?.v);
        
        if (isActionAssigned) {
          assignedCount++;
          
          if (dueDateCol) {
            const dueDateCellAddr = `${dueDateCol}${row + 1}`;
            const dueDateCell = worksheet[dueDateCellAddr];
            const dueDate = parseDate(dueDateCell?.w || dueDateCell?.v);
            
            if (dueDate && dueDate < today) {
              overdueCount++;
            }
          }
        }
      }

      const assignedPct = totalRows > 0 ? (assignedCount / totalRows) * 100 : 0;
      const overduePctOfAssigned = assignedCount > 0 ? (overdueCount / assignedCount) * 100 : 0;

      const summary = `Total: ${totalRows} | Assigned: ${assignedCount} (${assignedPct.toFixed(0)}%) | Overdue: ${overdueCount} (${overduePctOfAssigned.toFixed(0)}%)`;

      const analysisResult: AnalysisResult = {
        total_rows: totalRows,
        assigned_count: assignedCount,
        assigned_pct: Math.round(assignedPct * 10) / 10,
        overdue_count: overdueCount,
        overdue_pct_of_assigned: Math.round(overduePctOfAssigned * 10) / 10,
        today_iso: today.toISOString().split('T')[0],
        timezone: 'Asia/Kolkata',
        columns_used: {
          action: actionCol,
          due_date: dueDateCol || 'Not found'
        },
        notes,
        summary
      };

      setResult(analysisResult);
      toast({
        title: "Analysis Complete",
        description: `Analyzed ${totalRows} rows successfully`,
      });

    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to analyze Excel file';
      setError(errorMessage);
      toast({
        title: "Analysis Failed",
        description: errorMessage,
        variant: "destructive",
      });
    } finally {
      setIsAnalyzing(false);
    }
  }, [toast]);

  const handleDrag = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    const files = Array.from(e.dataTransfer.files);
    const excelFile = files.find(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (excelFile) {
      analyzeExcel(excelFile);
    } else {
      toast({
        title: "Invalid File Type",
        description: "Please upload an Excel file (.xlsx or .xls)",
        variant: "destructive",
      });
    }
  }, [analyzeExcel, toast]);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      analyzeExcel(file);
    }
  }, [analyzeExcel]);

  return (
    <div className="min-h-screen bg-background p-6">
      <div className="max-w-4xl mx-auto space-y-6">
        <div className="text-center space-y-4">
          <h1 className="text-4xl font-bold text-foreground">Excel Action Analyzer</h1>
          <p className="text-xl text-muted-foreground">
            Analyze Action and Due Date columns to track completion and overdue items
          </p>
        </div>

        <Card className="border-2 border-dashed border-muted-foreground/25 hover:border-primary/50 transition-colors">
          <CardContent className="p-8">
            <div
              className={`relative border-2 border-dashed rounded-lg p-12 text-center transition-all ${
                dragActive 
                  ? 'border-primary bg-primary/5 scale-105' 
                  : 'border-muted-foreground/25 hover:border-primary/50'
              }`}
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileInput}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                disabled={isAnalyzing}
              />
              
              <div className="space-y-4">
                <div className="mx-auto w-16 h-16 bg-primary/10 rounded-full flex items-center justify-center">
                  {isAnalyzing ? (
                    <div className="animate-spin w-8 h-8 border-2 border-primary border-t-transparent rounded-full" />
                  ) : (
                    <Upload className="w-8 h-8 text-primary" />
                  )}
                </div>
                
                <div>
                  <h3 className="text-xl font-semibold text-foreground">
                    {isAnalyzing ? 'Analyzing...' : 'Drop Excel file here'}
                  </h3>
                  <p className="text-muted-foreground mt-2">
                    Or click to browse for .xlsx or .xls files
                  </p>
                </div>
                
                <Button variant="outline" disabled={isAnalyzing}>
                  <FileSpreadsheet className="w-4 h-4 mr-2" />
                  Choose File
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>

        {error && (
          <Card className="border-destructive/50 bg-destructive/5">
            <CardContent className="p-6">
              <div className="flex items-center space-x-3">
                <AlertCircle className="w-5 h-5 text-destructive" />
                <div>
                  <h3 className="font-semibold text-destructive">Analysis Error</h3>
                  <p className="text-sm text-destructive/80 mt-1">{error}</p>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

        {result && (
          <div className="space-y-6">
            <Card className="bg-gradient-to-r from-primary/5 to-success/5 border-primary/20">
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <CheckCircle className="w-6 h-6 text-success" />
                  <span>Analysis Summary</span>
                </CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-lg font-mono bg-card p-4 rounded-lg border">
                  {result.summary}
                </p>
              </CardContent>
            </Card>

            <div className="grid md:grid-cols-3 gap-6">
              <Card>
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg">Total Items</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold text-primary">
                    {result.total_rows}
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg">Assigned</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold text-success">
                    {result.assigned_count}
                  </div>
                  <div className="text-sm text-muted-foreground">
                    {result.assigned_pct}% of total
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg">Overdue</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold text-destructive">
                    {result.overdue_count}
                  </div>
                  <div className="text-sm text-muted-foreground">
                    {result.overdue_pct_of_assigned}% of assigned
                  </div>
                </CardContent>
              </Card>
            </div>

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <Calendar className="w-5 h-5" />
                  <span>Analysis Details</span>
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid md:grid-cols-2 gap-4 text-sm">
                  <div>
                    <span className="font-medium">Analysis Date:</span> {result.today_iso}
                  </div>
                  <div>
                    <span className="font-medium">Timezone:</span> {result.timezone}
                  </div>
                  <div>
                    <span className="font-medium">Action Column:</span> {result.columns_used.action}
                  </div>
                  <div>
                    <span className="font-medium">Due Date Column:</span> {result.columns_used.due_date}
                  </div>
                </div>
                
                {result.notes.length > 0 && (
                  <div className="bg-warning/10 border border-warning/20 rounded-lg p-4">
                    <h4 className="font-medium text-warning-foreground mb-2">Notes:</h4>
                    <ul className="list-disc list-inside space-y-1 text-sm text-warning-foreground/80">
                      {result.notes.map((note, index) => (
                        <li key={index}>{note}</li>
                      ))}
                    </ul>
                  </div>
                )}
                
                <details className="bg-muted/50 rounded-lg p-4">
                  <summary className="font-medium cursor-pointer">View JSON Output</summary>
                  <pre className="mt-3 text-xs bg-card p-3 rounded border overflow-x-auto">
                    {JSON.stringify(result, null, 2)}
                  </pre>
                </details>
              </CardContent>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelAnalyzer;