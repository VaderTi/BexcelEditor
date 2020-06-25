using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace BexcelEditor.Definitions
{
    internal static class BexcelConverter
    {
        public static Bexcel Read(string inputPath)
        {
            try
            {
                var bexcel = new Bexcel();

                using (var bReader = new BinaryReader(File.Open(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read), Encoding.Unicode))
                {
                    var sheetCount = bReader.ReadInt32();

                    for (var i = 0; i < sheetCount; i++)
                    {
                        var tableName = ReadString(bReader);
                        var tableType = bReader.ReadInt32();

                        bexcel.Sheets.Add(new Bexcel.Sheet
                        {
                            Name = tableName,
                            Type = tableType
                        });
                    }

                    sheetCount = bReader.ReadInt32();
                    for (var j = 0; j < sheetCount; j++)
                    {
                        var sheetName = ReadString(bReader);
                        var currentSheet = bexcel.Sheets.First(x => x.Name == sheetName);

                        currentSheet.Index1 = bReader.ReadInt32();
                        currentSheet.Index2 = bReader.ReadInt32();

                        var columnCount = bReader.ReadInt32();
                        for (var k = 0; k < columnCount; k++)
                        {
                            currentSheet.Columns.Add(new Bexcel.Column
                            {
                                //Index = k,
                                Name = ReadString(bReader)
                            });
                        }

                        var rowCount = bReader.ReadInt32();
                        for (var l = 0; l < rowCount; l++)
                        {
                            var row = new Bexcel.Row
                            {
                                //Index = l,
                                Index1 = bReader.ReadInt32(),
                                Index2 = bReader.ReadInt32()
                            };

                            var cellCount = bReader.ReadInt32();

                            for (var m = 0; m < cellCount; m++)
                            {
                                row.Cells.Add(new Bexcel.Cell
                                {
                                    Index = m,
                                    Name = ReadString(bReader)
                                });
                            }

                            currentSheet.Rows.Add(row);
                        }

                        var columns = bReader.ReadInt32();
                        for (var n = 0; n < columns; n++)
                        {
                            currentSheet.Unknown1.Add(new Bexcel.Unknown
                            {
                                Index = n,
                                Text = ReadString(bReader),
                                Number = bReader.ReadInt32()
                            });
                        }

                        var rowCount2 = bReader.ReadInt32();
                        for (var num8 = 0; num8 < rowCount2; num8++)
                        {
                            currentSheet.TableDetails.Add(new Bexcel.Unknown
                            {
                                Index = num8,
                                Text = ReadString(bReader),
                                Number = bReader.ReadInt32()
                            });
                        }
                    }

                    bexcel.FileEnding = ReadString(bReader);
                }

                Debug.WriteLine("Read");

                return bexcel;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            return null;
        }

        public static void Save(Bexcel bexcel, string outputPath)
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            using (var fs = new FileStream(outputPath, FileMode.CreateNew))
            {
                using (var bw = new BinaryWriter(fs, Encoding.Unicode))
                {
                    bw.Write(bexcel.Sheets.Count); // sheetCount - Int32

                    foreach (var sheet in bexcel.Sheets.OrderBy(x => x.Type))
                    {
                        bw.Write((long)sheet.Name.Length);
                        bw.Write(WriteString(sheet.Name));
                        bw.Write(sheet.Type);
                    }

                    bw.Write(bexcel.Sheets.Count); // sheetCount - Int32

                    
                    foreach (var sheet in bexcel.Sheets.OrderBy(x => x.Type))
                    {
                        bw.Write((long)sheet.Name.Length);
                        bw.Write(WriteString(sheet.Name));

                        bw.Write(sheet.Index1);
                        bw.Write(sheet.Index2);

                        bw.Write(sheet.Columns.Count);

                        foreach (var column in sheet.Columns)
                        {
                            bw.Write((long)column.Name.Length);
                            bw.Write(WriteString(column.Name));
                        }

                        bw.Write(sheet.Rows.Count); // row count
                        foreach (var row in sheet.Rows)
                        {
                            bw.Write(row.Index1);
                            bw.Write(row.Index2);
                            bw.Write(row.Cells.Count);

                            foreach (var cells in row.Cells)
                            {
                                bw.Write((long)cells.Name.Length);
                                bw.Write(WriteString(cells.Name));
                            }
                        }

                        bw.Write(sheet.Unknown1.Count);
                        foreach (var unk in sheet.Unknown1)
                        {
                            bw.Write((long)unk.Text.Length);
                            bw.Write(WriteString(unk.Text));
                            bw.Write(unk.Number);
                        }

                        bw.Write(sheet.TableDetails.Count);
                        foreach (var unk in sheet.TableDetails)
                        {
                            bw.Write((long)unk.Text.Length);
                            bw.Write(WriteString(unk.Text));
                            bw.Write(unk.Number);
                        }
                    }

                    bw.Write((long)bexcel.FileEnding.Length);
                    bw.Write(WriteString(bexcel.FileEnding));
                }
            }

            Debug.WriteLine("Saved");
            Debug.WriteLine(bexcel.FileEnding);
        }

        public static DataTable ToDataTable(Bexcel.Sheet sheet)
        {
            var dt = new DataTable();

            foreach (var column in sheet.Columns)
            {
                dt.Columns.Add(column.Name);
            }

            foreach (var row in sheet.Rows)
            {
                var dr = dt.NewRow();
                var i = 0;
                foreach (var cell in row.Cells)
                {
                    dr[i] = cell.Name;
                    i++;
                }
                dt.Rows.Add(dr);
            }

            dt.TableName = sheet.Name;

            return dt;
        }

        private static string ReadString(BinaryReader r)
        {
            return Encoding.Unicode.GetString(r.ReadBytes((int)r.ReadInt64() * 2));
        }

        private static byte[] WriteString(string s)
        {
            return Encoding.Unicode.GetBytes(s);
        }
    }
}