using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Sys = System.Windows.Forms;
using System.Threading;

namespace Gra_w_zbijanego
{
    public partial class Ribbon1
    {
        private static int paddle = 29;
        private static int trap = 43;
        private static int point = 3;
        private static int white = 2;
        private Excel.Range rCenter;
        private static int iCenter = 50;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void GameNewGame_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet Gra;
            Excel.Workbook w_ThisWorkbook = Globals.ThisAddIn.GetActiveWorkbook();
            
            Boolean AgreeToCreate = false;

            if( WorksheetExists(w_ThisWorkbook, "Game"))
            {
                String msg = "Worksheet 'Game' already exists in this workbook. \nDo you want to " +
                    "create game template in this sheet?";
                String heading = "'Game' worksheet";


                Sys.DialogResult result = Sys.MessageBox.Show(msg, heading, Sys.MessageBoxButtons.YesNo);

                if(result== Sys.DialogResult.Yes)
                {
                    AgreeToCreate = true;
                }
                Gra = w_ThisWorkbook.Worksheets["Game"];
            }else
            {

                Gra = w_ThisWorkbook.Worksheets.Add();
                Gra.Name = "Game";
                AgreeToCreate = true;
            }

            if (AgreeToCreate)
            {
                CreateTemplate(Gra);
            }
 
        }

        private Boolean WorksheetExists( Excel.Workbook wbk, String name){

            Boolean b;

            try
            {
               b= (int)wbk.Worksheets[name].Index > 0;
            }catch(Exception)
            {
                b = false;
            }

            return b;
        }

        private void CreateTemplate(Excel.Worksheet wks){

            wks.Cells.ClearContents();
            wks.Cells.Interior.ColorIndex = 1;
            wks.Range["A1:Y25"].Interior.ColorIndex = white;
            wks.Cells.ColumnWidth = 2;
            wks.Cells.RowHeight = 12;
            
        }

 

        private void Start_Click(object sender, RibbonControlEventArgs e)
        {
            

            bool lose = false;
            int speed = 130;
            int score = 0;
            Excel.Worksheet w_Game = Globals.ThisAddIn.GetActiveWorkbook().Worksheets["Game"];

            for (int k = 1; k < 10; k++)
            {
                int leftPadle = 11;
                int rightPadle = 13;
                int xPoint = 13;
                int yPoint = 23;
                int xMove = 1;
                int yMove = -1;
                CreateGame(w_Game);
                rCenter = w_Game.Cells[26, 50];
                rCenter.Select();
                for (int j =1; j<400*k ; j++ )
                {
                    score = score + 1;
                    w_Game.Cells[yPoint, xPoint].Interior.ColorIndex = 2;
                    for (int i = 1; i < 6; i++)
                    {
                        MovePaddle(w_Game, ref leftPadle, ref rightPadle);
                    }
                    //Odbicie od bocznych scian

                    if (xPoint == 25 || xPoint == 1)
                    {
                        xMove = -xMove;
                    }
                    else if (w_Game.Cells[yPoint, xPoint - 1].Interior.ColorIndex == trap)
                    {
                        xMove = -xMove;
                        w_Game.Cells[yPoint, xPoint - 1].Interior.ColorIndex = white;
                    }
                    else if (w_Game.Cells[yPoint, xPoint + 1].Interior.ColorIndex == trap)
                    {
                        xMove = -xMove;
                        w_Game.Cells[yPoint, xPoint + 1].Interior.ColorIndex = white;
                    }
                    else if (xMove == -1 && yMove == 1 && w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex == trap)
                    {
                        w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex = white;
                        xMove = 1;
                        yMove = -1;
                    }
                    else if (xMove == 1 && yMove == 1 && w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex == trap)
                    {
                        w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex = white;
                        xMove = -1;
                        yMove = -1;
                    }

                    for (int i = 1; i < 6; i++)
                    {
                        MovePaddle(w_Game, ref leftPadle, ref rightPadle);
                    }

                    if (yPoint == 1 || w_Game.Cells[yPoint + 1, xPoint].Interior.ColorIndex == paddle)
                    {
                        yMove = -yMove;
                    }
                    else if (w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex == paddle)
                    {
                        yMove = -yMove;
                        xMove = -xMove;
                    }
                    else if (w_Game.Cells[yPoint + 1, xPoint].Interior.ColorIndex == trap)
                    {
                        yMove = -yMove;
                        w_Game.Cells[yPoint + 1, xPoint].Interior.ColorIndex = white;
                    }
                    else if (w_Game.Cells[yPoint - 1, xPoint].Interior.ColorIndex == trap)
                    {
                        yMove = -yMove;
                        w_Game.Cells[yPoint - 1, xPoint].Interior.ColorIndex = white;
                    }
                    else if (xMove == -1 && yMove == -1 && w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex == trap)
                    {
                        w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex = white;
                        xMove = 1;
                        yMove = 1;
                    }
                    else if (xMove == 1 && yMove == -1 && w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex == trap)
                    {
                        w_Game.Cells[yPoint + yMove, xPoint + xMove].Interior.ColorIndex = white;
                        xMove = -1;
                        yMove = 1;
                    }

                    for (int i = 1; i < 6; i++)
                    {
                        MovePaddle(w_Game, ref leftPadle, ref rightPadle);
                    }



                    yPoint = yPoint + yMove;
                    xPoint = xPoint + xMove;

                    w_Game.Cells[yPoint, xPoint].Interior.ColorIndex = point;

                    if (yPoint == 25)
                    {
                        lose = true;
                        break;
                    }
                    for (int i = 1; i < 6; i++)
                    {
                        MovePaddle(w_Game, ref leftPadle, ref rightPadle);
                    }

                    Thread.Sleep(speed);

                    for (int i = 1; i < 6; i++)
                    {
                        MovePaddle(w_Game, ref leftPadle, ref rightPadle);
                    }
                }
                if (lose) { break; }
                speed = speed - 10;
                score = score + 100;
            }

            Sys.MessageBox.Show("You lose! \nScore: " + score +"!");
        }

        private void CreateGame(Excel.Worksheet wbk)
        {
            wbk.Range["A1:Y25"].Interior.ColorIndex = white;
            wbk.Cells[25, 11].Resize[1, 3].Interior.ColorIndex =paddle;
            wbk.Range["C3"].Resize[10, 21].Interior.ColorIndex = trap;
            wbk.Cells[23, 13].Interior.ColorIndex = point;
            
        }

        private void MovePaddle( Excel.Worksheet wks, ref int left, ref int right)
        {
            Sys.Application.DoEvents();
            int i_column = Globals.ThisAddIn.GetActiveCell().Column;
            if (i_column > iCenter && right != 25) 
            {
                wks.Cells[25, left].Interior.ColorIndex = white;
                wks.Cells[25, right + 1].Interior.ColorIndex = paddle;

                left = left + 1;
                right = right + 1;

            }else if(i_column < iCenter && left!= 1)
            {
                wks.Cells[25, right].Interior.ColorIndex = white;
                wks.Cells[25, left - 1].Interior.ColorIndex = paddle;

                left = left - 1;
                right = right - 1;
            }
            rCenter.Select();
        }




    }
}
