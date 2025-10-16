import ExcelJS from 'exceljs';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  WidthType,
  ImageRun,
  BorderStyle
} from 'docx';
import { saveAs } from 'file-saver';

// ====== GRADE & UTILITY FUNCTIONS ======

export const calculateGrade = (score) => {
  if (!score && score !== 0) return '-';
  const numScore = parseFloat(score);
  if (numScore >= 75) return 'A';
  if (numScore >= 70) return 'A-';
  if (numScore >= 65) return 'B+';
  if (numScore >= 60) return 'B';
  if (numScore >= 55) return 'B-';
  if (numScore >= 50) return 'C+';
  if (numScore >= 45) return 'C';
  if (numScore >= 40) return 'C-';
  if (numScore >= 35) return 'D+';
  if (numScore >= 30) return 'D';
  if (numScore >= 25) return 'D-';
  return 'E';
};

export const calculateTotalGrade = (totalScore, maxPossible = 1100) => {
  if (!totalScore && totalScore !== 0) return '-';
  const percentage = (totalScore / maxPossible) * 100;
  if (percentage >= 78) return 'A';
  if (percentage >= 75) return 'A-';
  if (percentage >= 70) return 'B+';
  if (percentage >= 65) return 'B';
  if (percentage >= 55) return 'B-';
  if (percentage >= 48) return 'C+';
  if (percentage >= 40) return 'C';
  if (percentage >= 35) return 'C-';
  if (percentage >= 30) return 'D+';
  if (percentage >= 25) return 'D';
  if (percentage >= 20) return 'D-';
  return 'E';
};

export const formatScoreWithGrade = (score) => {
  if (!score && score !== 0) return '-';
  const grade = calculateGrade(score);
  return `${Math.round(score)}${grade}`;
};

export const getUniqueSubjects = (students) => {
  const subjects = new Set();
  students.forEach(student => {
    (student.subject_scores || []).forEach(score => {
      if (score.subject_name) {
        subjects.add(score.subject_name);
      }
    });
  });
  return Array.from(subjects).sort();
};

// ====== DEFAULT TEACHERS MAP ======
const DEFAULT_TEACHERS = {
  'BUSINESS STUDIES': 'OJWANG W.',
  'CRE': 'MILKA OYIER',
  'PHYSICS': 'ODWAR J',
  'MATHEMATICS': 'ODWAR J.',
  'CHEMISTRY': 'KENNEDY O.',
  'HISTORY': 'PAUL O',
  'AGRICULTURE': 'BRIAN O.',
  'KISWAHILI': 'OJWANG W',
  'ENGLISH': 'BRIAN O',
  'BIOLOGY': 'KENNEDY O.',
  'GEOGRAPHY': 'ODHIAMBO C'
};

const getTeacherForSubject = (subjectName) => {
  if (!subjectName) return 'N/A';
  return DEFAULT_TEACHERS[subjectName.toUpperCase().trim()] || 'N/A';
};

// ====== EXCEL EXPORT ======

const calculateSubjectMean = (students, subject) => {
  const scores = students.map(student => {
    const subjectScore = student.subject_scores?.find(s => s.subject_name === subject);
    return subjectScore ? parseFloat(subjectScore.score) || 0 : 0;
  });
  const sum = scores.reduce((total, score) => total + score, 0);
  return scores.length > 0 ? sum / scores.length/ 12 : 0;
};

const calculateClassMean = (students) => {
  const totalScores = students.map(student =>
    student.total_score ||
    student.subject_scores?.reduce((sum, score) => sum + (parseFloat(score.score) || 0), 0) || 0
  );
  const sum = totalScores.reduce((total, score) => total + score, 0);
  return totalScores.length > 0 ? sum / totalScores.length / 12: 0;
};

const calculateSubjectRankings = (students, subjects) => {
  return subjects.map(subject => {
    const scores = students
      .map(student => {
        const scoreObj = student.subject_scores?.find(s => s.subject_name === subject);
        return scoreObj ? parseFloat(scoreObj.score) || 0 : 0;
      })
      .filter(score => score > 0);
    const mean = scores.length > 0 ? scores.reduce((sum, score) => sum + score, 0) / scores.length / 12 : 0;
    const grade = calculateGrade(mean);
    return {
      subject,
      mean,
      grade,
      teacher: getTeacherForSubject(subject)
    };
  })
  .sort((a, b) => b.mean - a.mean)
  .map((stat, index) => ({ ...stat, position: index + 1 }));
};

const getSubjectAbbreviation = (subjectName) => {
  const abbreviations = {
    'ENGLISH': 'ENG',
    'KISWAHILI': 'KIS',
    'MATHEMATICS': 'MAT',
    'CHEMISTRY': 'CHEM',
    'BIOLOGY': 'BIO',
    'PHYSICS': 'PHY',
    'HISTORY': 'HIST',
    'GEOGRAPHY': 'GEO',
    'CRE': 'CRE',
    'AGRICULTURE': 'AGR',
    'BUSINESS': 'BST'
  };
  const upperSubject = subjectName.toUpperCase();
  return abbreviations[upperSubject] || subjectName.substring(0, 4).toUpperCase();
};

const getGradeRemarks = (grade) => {
  const remarks = {
    'A': 'Excellent - Outstanding performance',
    'A-': 'Excellent - Very strong performance',
    'B+': 'Very Good - Strong performance',
    'B': 'Good - Above average performance',
    'B-': 'Good - Good performance',
    'C+': 'Average - Satisfactory performance',
    'C': 'Below Average - Needs improvement',
    'C-': 'Below Average - Requires attention',
    'D+': 'Weak - Significant improvement needed',
    'D': 'Weak - Much improvement required',
    'D-': 'Weak - Serious attention needed',
    'E': 'Poor - Immediate intervention required'
  };
  return remarks[grade] || 'Performance assessment pending';
};

export const exportResultsToExcel = async (students, filters = {}, filterOptions = {}) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Results');
    const subjects = getUniqueSubjects(students);

    const baseColumns = [
      { header: 'ADM.NO', key: 'admission_no', width: 8 },
      { header: 'NAME', key: 'name', width: 30 }
    ];
    const subjectColumns = subjects.map(subject => ({
      header: getSubjectAbbreviation(subject),
      key: `subject_${subject}`,
      width: 8
    }));
    const summaryColumns = [
      { header: 'TT MARKS', key: 'total_marks', width: 10 },
      { header: 'GRADE', key: 'total_grade', width: 8 },
      { header: 'C RANK', key: 'class_rank', width: 8 }
    ];
    const allColumns = [...baseColumns, ...subjectColumns, ...summaryColumns];
    worksheet.columns = allColumns;

    for (let i = 0; i < 4; i++) {
      worksheet.addRow([]);
    }

    const headerRow = worksheet.getRow(5);
    headerRow.values = allColumns.map(col => col.header);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: 'center' };
    headerRow.height = 25;

    let currentRow = 6;
    const sortedStudents = [...students].sort((a, b) => {
      const rankA = a.class_rank || Number.MAX_SAFE_INTEGER;
      const rankB = b.class_rank || Number.MAX_SAFE_INTEGER;
      return rankA - rankB;
    });

    sortedStudents.forEach((student) => {
      const rowData = {
        admission_no: student.admission_number || student.admission_no,
        name: student.fullname || student.name
      };

      subjects.forEach(subject => {
        const subjectScore = student.subject_scores?.find(score => score.subject_name === subject);
        const score = subjectScore ? subjectScore.score : '-';
        rowData[`subject_${subject}`] = score !== '-' ? formatScoreWithGrade(score) : '-';
      });

      const totalMarks = student.total_score ||
        student.subject_scores?.reduce((sum, score) => sum + (parseFloat(score.score) || 0), 0) || 0;
      rowData.total_marks = Math.round(totalMarks);
      rowData.total_grade = calculateTotalGrade(totalMarks);
      rowData.class_rank = student.class_rank || '';

      const row = worksheet.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        if (colNumber !== 2) {
          cell.alignment = { horizontal: 'center' };
        }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    const statsStartRow = currentRow + sortedStudents.length + 2;
    worksheet.mergeCells(`A${statsStartRow}:B${statsStartRow}`);
    worksheet.getCell(`A${statsStartRow}`).value = 'MEAN';
    worksheet.getCell(`A${statsStartRow}`).font = { bold: true };

    subjects.forEach((subject, index) => {
      const colLetter = String.fromCharCode(67 + index);
      const meanScore = calculateSubjectMean(sortedStudents, subject);
      worksheet.getCell(`${colLetter}${statsStartRow}`).value = meanScore.toFixed(2);
      worksheet.getCell(`${colLetter}${statsStartRow}`).alignment = { horizontal: 'center' };
    });

    worksheet.addRow([]);
    worksheet.addRow([]);

    const ranksStartRow = statsStartRow + 3;
    worksheet.mergeCells(`A${ranksStartRow}:E${ranksStartRow}`);
    worksheet.getCell(`A${ranksStartRow}`).value = 'SUBJECT RANKS';
    worksheet.getCell(`A${ranksStartRow}`).font = { bold: true, size: 14 };

    const ranksHeaderRow = worksheet.getRow(ranksStartRow + 1);
    ranksHeaderRow.values = ['SUBJECT', 'MEAN', 'GRADE', 'TEACHER', 'POSITION'];
    ranksHeaderRow.font = { bold: true };
    ranksHeaderRow.alignment = { horizontal: 'center' };

    const subjectRanks = calculateSubjectRankings(sortedStudents, subjects);
    subjectRanks.forEach((rank) => {
      const row = worksheet.addRow([
        getSubjectAbbreviation(rank.subject),
        rank.mean.toFixed(2),
        rank.grade,
        rank.teacher,
        rank.position
      ]);
      row.alignment = { horizontal: 'center' };
    });

    const classMeanRow = ranksStartRow + subjectRanks.length + 2;
    worksheet.mergeCells(`A${classMeanRow}:D${classMeanRow}`);
    worksheet.getCell(`A${classMeanRow}`).value = 'CLASS MEAN';
    worksheet.getCell(`A${classMeanRow}`).font = { bold: true };
    const classMean = calculateClassMean(sortedStudents);
    worksheet.getCell(`E${classMeanRow}`).value = classMean.toFixed(2);
    worksheet.getCell(`E${classMeanRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`F${classMeanRow}`).value = calculateGrade(classMean);
    worksheet.getCell(`F${classMeanRow}`).alignment = { horizontal: 'center' };

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  } catch (error) {
    console.error('Error exporting to Excel:', error);
    throw new Error('Failed to generate Excel file: ' + error.message);
  }
};

// ====== IMAGE LOADER ======

export const loadImageAsBase64 = async (imageUrl) => {
  try {
    const response = await fetch(imageUrl);
    if (!response.ok) throw new Error(`Failed to fetch image: ${response.statusText}`);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const result = reader.result;
        if (typeof result === 'string') {
          resolve(result.split(',')[1]);
        } else {
          reject(new Error('Failed to read image as base64'));
        }
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error('Error loading image as base64:', error);
    return null;
  }
};

// ====== WORD EXPORT (with Navy + Teal) ======

const createInfoRow = (label, value) => {
  return new Paragraph({
    children: [
      new TextRun({ text: label, bold: true, size: 16 }),
      new TextRun({ text: "\t" + (value || 'N/A'), size: 16 }),
    ],
    spacing: { after: 120 },
  });
};

const getGradeColor = (grade) => {
  const colors = {
    'A': '2E8B57',   // Emerald Green
    'A-': '2E8B57',
    'B+': '28A79A',  // Teal
    'B': '28A79A',
    'B-': '28A79A',
    'C+': 'FFA726',  // Amber
    'C': 'FFA726',
    'C-': 'FFA726',
    'D+': 'EF5350',  // Soft Red
    'D': 'EF5350',
    'D-': 'EF5350',
    'E': '990000'    // Dark Red
  };
  return colors[grade] || '000000';
};

const createSubjectTable = (subjects) => {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [2500, 1200, 1200, 2000, 2000],
    rows: [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "SUBJECT", bold: true, size: 14, color: "FFFFFF" })], alignment: AlignmentType.CENTER })], shading: { fill: "0A2E5C" } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "SCORE", bold: true, size: 14, color: "FFFFFF" })], alignment: AlignmentType.CENTER })], shading: { fill: "0A2E5C" } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "GRADE", bold: true, size: 14, color: "FFFFFF" })], alignment: AlignmentType.CENTER })], shading: { fill: "0A2E5C" } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TEACHER", bold: true, size: 14, color: "FFFFFF" })], alignment: AlignmentType.CENTER })], shading: { fill: "0A2E5C" } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "REMARKS", bold: true, size: 14, color: "FFFFFF" })], alignment: AlignmentType.CENTER })], shading: { fill: "0A2E5C" } }),
        ],
      }),
      ...subjects.map((subject, index) => {
        const score = Math.round(subject.score || 0);
        const grade = subject.grade || calculateGrade(subject.score);
        const teacher = subject.teacher || getTeacherForSubject(subject.subject_name);
        const backgroundColor = index % 2 === 0 ? "F8F9FA" : "FFFFFF";
        return new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: subject.subject_name || subject.name, bold: true, size: 13 })], alignment: AlignmentType.LEFT })], shading: { fill: backgroundColor } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(score), bold: true, size: 13 })], alignment: AlignmentType.CENTER })], shading: { fill: backgroundColor } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: grade, bold: true, size: 13, color: getGradeColor(grade) })], alignment: AlignmentType.CENTER })], shading: { fill: backgroundColor } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: teacher, size: 13 })], alignment: AlignmentType.CENTER })], shading: { fill: backgroundColor } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: subject.remarks || getGradeRemarks(grade), size: 13 })], alignment: AlignmentType.CENTER })], shading: { fill: backgroundColor } }),
          ],
        });
      }),
    ],
  });
};

export const exportIndividualResultToWord = async (student, subjects = [], comments = {}, imageBase64 = null) => {
  try {
    // Inject teacher names if missing
    const enrichedSubjects = subjects.map(subj => ({
      ...subj,
      teacher: subj.teacher || getTeacherForSubject(subj.subject_name)
    }));

    let imageRun = null;
    if (imageBase64) {
      imageRun = new ImageRun({
        data: Uint8Array.from(atob(imageBase64), c => c.charCodeAt(0)),
        transformation: { width: 80, height: 80 },
      });
    }

    const totalMarks = student.total_score || enrichedSubjects.reduce((sum, s) => sum + (parseFloat(s.score) || 0), 0);
    const percentage = ((totalMarks / 1100) * 100).toFixed(1);

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          // Header with Logo
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2000, 8000],
            rows: [new TableRow({
              children: [
                new TableCell({
                  children: imageRun ? [new Paragraph({ children: [imageRun], alignment: AlignmentType.CENTER })] : [
                    new Paragraph({ children: [new TextRun({ text: "[LOGO]", size: 16 })], alignment: AlignmentType.CENTER })
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: "ST PETERS MAWENI GIRLS SECONDARY SCHOOL", bold: true, size: 22, allCaps: true })], alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "P.O. BOX 941-40400 SUNA MIGORI", size: 18 })], alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "Email: stpetersmaweni@gmail.com", size: 16 })], alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "KNOWLEDGE IS POWER", italics: true, size: 16 })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
              ],
            })],
          }),

          // Separator
          new Paragraph({
            children: [new TextRun({ text: "―".repeat(80), size: 16, color: "0A2E5C" })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 400 },
          }),

          // Report Title
          new Paragraph({
            children: [new TextRun({ text: "INDIVIDUAL ACADEMIC PERFORMANCE REPORT", bold: true, size: 24, allCaps: true, color: "0A2E5C" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 },
          }),

          // Info Grid
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [5000, 5000],
            rows: [new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: "STUDENT INFORMATION", bold: true, size: 18, color: "0A2E5C" })], spacing: { after: 200 } }),
                    createInfoRow("Full Name:", student.fullname || student.name),
                    createInfoRow("Admission No:", student.admission_number || student.admission_no),
                    createInfoRow("Class:", `${student.class_name || student.form} ${student.stream_name || student.stream || ''}`),
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: "ACADEMIC SUMMARY", bold: true, size: 18, color: "0A2E5C" })], spacing: { after: 200 } }),
                    createInfoRow("Total Marks:", Math.round(totalMarks).toString()),
                    createInfoRow("Percentage:", `${percentage}%`),
                    createInfoRow("Overall Grade:", student.overall_grade || calculateTotalGrade(totalMarks)),
                    createInfoRow("Class Rank:", student.class_rank ? `#${student.class_rank}` : 'N/A'),
                    createInfoRow("Stream Rank:", student.stream_rank ? `#${student.stream_rank}` : 'N/A'),
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
              ],
            })],
          }),

          new Paragraph({ text: "", spacing: { after: 400 } }),

          // Subject Table
          new Paragraph({
            children: [new TextRun({ text: "SUBJECT PERFORMANCE ANALYSIS", bold: true, size: 18, color: "0A2E5C" })],
            spacing: { before: 200, after: 300 },
          }),
          enrichedSubjects.length > 0 ? createSubjectTable(enrichedSubjects) :
            new Paragraph({ text: "No subject performance data available", alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 } }),

          new Paragraph({ text: "", spacing: { after: 400 } }),

          // Comments
          new Paragraph({
            children: [new TextRun({ text: "OFFICIAL COMMENTS", bold: true, size: 18, color: "0A2E5C" })],
            spacing: { before: 200, after: 300 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Principal's Comment: ", bold: true, size: 16 }),
              new TextRun({ text: comments.principal || ".........................." }),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Class Teacher's Comment: ", bold: true, size: 16 }),
              new TextRun({ text: comments.class_teacher || "..........................", size: 16 }),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Fee Balance: ", bold: true, size: 16 }),
              new TextRun({
                text: "KSh " + (student.fee_balance || ".............."),
                size: 16,
                color: student.fee_balance > 0 ? "EF5350" : "2E8B57",
              }),
            ],
            spacing: { before: 400, after: 200 },
          }),

          // Signatures
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [5000, 5000],
            rows: [new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: "_________________________", size: 14 })], alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "Principal's Signature", bold: true, size: 14 })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: "_________________________", size: 14 })], alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "Class Teacher's Signature", bold: true, size: 14 })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
                  ],
                  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
                }),
              ],
            })],
          }),

          // Footer
          new Paragraph({
            children: [new TextRun({ text: "Generated by Leratech Academic System", size: 12, color: "666666" })],
            alignment: AlignmentType.RIGHT,
            spacing: { before: 600 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `Report generated on: ${new Date().toLocaleDateString()}`, size: 10, color: "999999" })],
            alignment: AlignmentType.RIGHT,
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    return blob;
  } catch (error) {
    console.error('Error generating Word document:', error);
    throw new Error('Failed to generate Word document: ' + error.message);
  }
};

// ====== HTML EXPORT (Modern UI) ======

export const exportIndividualResultAsHTML = async (student, subjects = [], comments = {}, logoUrl = null) => {
  const enrichedSubjects = subjects.map(subj => ({
    ...subj,
    teacher: subj.teacher || getTeacherForSubject(subj.subject_name)
  }));

  const totalMarks = student.total_score || enrichedSubjects.reduce((sum, s) => sum + (parseFloat(s.score) || 0), 0);
  const percentage = ((totalMarks / 1100) * 100).toFixed(1);
  const overallGrade = student.overall_grade || calculateTotalGrade(totalMarks);
  const gradeClass = overallGrade.startsWith('A') ? 'A' :
                     overallGrade.startsWith('B') ? 'B' :
                     overallGrade.startsWith('C') ? 'C' : 'D';

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>Individual Academic Report - ${student.fullname || student.name}</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        body { 
          font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
          margin: 0;
          padding: 40px;
          line-height: 1.6;
          color: #333;
          background: #F8F9FA;
        }
        .report-container {
          max-width: 1000px;
          margin: 0 auto;
          background: white;
          border-radius: 12px;
          box-shadow: 0 6px 20px rgba(10, 46, 92, 0.12);
          overflow: hidden;
        }
        .header {
          background: linear-gradient(135deg, #0A2E5C 0%, #1A3F6D 100%);
          color: white;
          padding: 30px;
          text-align: center;
        }
        .school-name {
          font-size: 24px;
          font-weight: 700;
          margin-bottom: 8px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .report-title {
          font-size: 28px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 1.5px;
          margin: 15px 0;
          color: #28A79A;
        }
        .section-title {
          font-size: 20px;
          font-weight: 700;
          color: #0A2E5C;
          margin-bottom: 20px;
          padding-bottom: 10px;
          border-bottom: 2px solid #E9ECEF;
        }
        .info-card {
          background: #F8F9FA;
          padding: 25px;
          border-radius: 10px;
          border-left: 4px solid #0A2E5C;
        }
        .info-card h3 {
          margin: 0 0 15px 0;
          color: #0A2E5C;
          font-size: 16px;
          font-weight: 700;
        }
        .grade-A { color: #2E8B57; font-weight: 700; }
        .grade-B { color: #28A79A; font-weight: 700; }
        .grade-C { color: #FFA726; font-weight: 700; }
        .grade-D { color: #EF5350; font-weight: 700; }

        .subject-table {
          width: 100%;
          border-collapse: collapse;
          margin: 20px 0;
          border-radius: 10px;
          overflow: hidden;
          box-shadow: 0 2px 8px rgba(10, 46, 92, 0.08);
        }
        .subject-table th {
          background: #0A2E5C;
          color: white;
          padding: 14px 16px;
          text-align: center;
          font-weight: 600;
          font-size: 14px;
        }
        .subject-table td {
          padding: 12px 16px;
          border-bottom: 1px solid #E9ECEF;
          text-align: center;
        }
        .subject-table tr:nth-child(even) {
          background: #FBFCFD;
        }
        .comments-section {
          background: #F8F9FA;
          padding: 25px;
          border-radius: 10px;
          margin: 30px 0;
        }
        .comment-label {
          font-weight: 700;
          color: #0A2E5C;
          margin-bottom: 8px;
          display: block;
        }
        .signature-section {
          margin-top: 50px;
          padding-top: 30px;
          border-top: 2px solid #E9ECEF;
        }
        .footer {
          margin-top: 40px;
          padding-top: 20px;
          border-top: 1px solid #E9ECEF;
          color: #6C757D;
          font-size: 12px;
          text-align: center;
        }
        @media (max-width: 768px) {
          body { padding: 20px; }
          .info-grid { grid-template-columns: 1fr; }
        }
      </style>
    </head>
    <body>
      <div class="report-container">
        <div class="header">
          <div class="school-name">St Peters Maweni Girls Secondary School</div>
          <div class="report-title">Individual Academic Performance Report</div>
        </div>

        <div class="content" style="padding: 40px;">
          <div class="info-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 30px;">
            <div class="info-card">
              <h3>STUDENT INFORMATION</h3>
              <div><strong>Full Name:</strong> ${student.fullname || student.name}</div>
              <div><strong>Admission No:</strong> ${student.admission_number || student.admission_no}</div>
              <div><strong>Class:</strong> ${student.class_name || student.form} ${student.stream_name || student.stream || ''}</div>
            </div>
            <div class="info-card">
              <h3>ACADEMIC SUMMARY</h3>
              <div><strong>Total Marks:</strong> ${Math.round(totalMarks)}</div>
              <div><strong>Percentage:</strong> ${percentage}%</div>
              <div><strong>Overall Grade:</strong> <span class="grade-${gradeClass}">${overallGrade}</span></div>
              <div><strong>Class Rank:</strong> ${student.class_rank ? '#' + student.class_rank : 'N/A'}</div>
              <div><strong>Stream Rank:</strong> ${student.stream_rank ? '#' + student.stream_rank : 'N/A'}</div>
            </div>
          </div>

          <div class="section">
            <div class="section-title">Subject Performance Analysis</div>
            ${enrichedSubjects.length > 0 ? `
              <table class="subject-table">
                <thead>
                  <tr>
                    <th>Subject</th>
                    <th>Score</th>
                    <th>Grade</th>
                    <th>Teacher</th>
                    <th>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  ${enrichedSubjects.map(subject => {
                    const score = Math.round(subject.score || 0);
                    const grade = subject.grade || calculateGrade(subject.score);
                    const gClass = grade.startsWith('A') ? 'A' :
                                   grade.startsWith('B') ? 'B' :
                                   grade.startsWith('C') ? 'C' : 'D';
                    return `
                      <tr>
                        <td><strong>${subject.subject_name || subject.name}</strong></td>
                        <td><strong>${score}</strong></td>
                        <td class="grade-${gClass}">${grade}</td>
                        <td>${subject.teacher || 'N/A'}</td>
                        <td>${subject.remarks || getGradeRemarks(grade)}</td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            ` : '<p style="text-align: center; color: #666;">No subject data available</p>'}
          </div>

          <div class="comments-section">
            <div class="section-title">Official Comments</div>
            <div>
              <span class="comment-label">Principal's Comment:</span>
              <div>${comments.principal || "Good performance. Maintain consistency and focus on continuous improvement."}</div>
            </div>
            <div style="margin-top: 15px;">
              <span class="comment-label">Class Teacher's Comment:</span>
              <div>${comments.class_teacher || "Shows great dedication and consistent improvement in academic performance."}</div>
            </div>
            <div style="margin-top: 15px;">
              <span class="comment-label">Fee Balance:</span>
              <div style="color: ${student.fee_balance > 0 ? '#EF5350' : '#2E8B57'};">
                KSh ${student.fee_balance || "0.00"}
              </div>
            </div>
          </div>

          <div class="signature-section">
            <div style="display: flex; justify-content: space-between; gap: 40px;">
              <div style="text-align: center; flex: 1;">
                <div style="border-bottom: 1px solid #333; padding-bottom: 5px; margin-bottom: 8px;">&nbsp;</div>
                <div>Principal's Signature</div>
              </div>
              <div style="text-align: center; flex: 1;">
                <div style="border-bottom: 1px solid #333; padding-bottom: 5px; margin-bottom: 8px;">&nbsp;</div>
                <div>Class Teacher's Signature</div>
              </div>
            </div>
          </div>

          <div class="footer">
            <div>Generated by Leratech Academic System</div>
            <div>Report generated on: ${new Date().toLocaleDateString()}</div>
          </div>
        </div>
      </div>
    </body>
    </html>
  `;

  return new Blob([htmlContent], { type: 'text/html' });
};

// ====== OTHER UTILS ======

export const downloadBlob = (blob, filename) => {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
  document.body.removeChild(a);
};

export const generateFilename = (filters, filterOptions, student = null) => {
  const term = filterOptions.terms?.find(t => t.term_id?.toString() === filters.term_id);
  const className = filterOptions.classes?.find(c => c.class_id?.toString() === filters.class_id)?.class_name;
  const streamName = filterOptions.streams?.find(s => s.stream_id?.toString() === filters.stream_id)?.stream_name;
  if (student) {
    return `Academic_Report_${student.admission_number}_${(student.fullname || student.name).replace(/\s+/g, '_')}.docx`;
  } else {
    const streamPart = streamName ? `_${streamName}` : '_all_streams';
    return `MAWENI_Results_${className}${streamPart}_${term?.academic_year || ''}.xlsx`;
  }
};

export default {
  exportResultsToExcel,
  exportIndividualResultToWord,
  exportIndividualResultAsHTML,
  downloadBlob,
  generateFilename,
  calculateGrade,
  calculateTotalGrade,
  formatScoreWithGrade,
  loadImageAsBase64
};
