using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Windows.Forms;
using P=DocumentFormat.OpenXml.Presentation;
using D=DocumentFormat.OpenXml.Drawing;
using System.Diagnostics;

namespace UpdateOffice
{
    public class UpdatePowerPoint
    {
        public static void UpdatePpt()
        {
            //从现有的ppt中复制出来新的一份
            string filePath = @"C:\Users\97470\Documents\1.pptx";
            string newFilePath = @"C:\Users\97470\Documents\" + DateTime.Now.ToString("MMdd") + ".pptx";
            File.Copy(filePath, newFilePath,true);

            //打开需要编辑的ppt 第二个参数判定该ppt是否允许修改
            PresentationDocument presentationDocument = PresentationDocument.Open(newFilePath, isEditable: true);

            //获取幻灯片的演示文稿
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            //获取该ppt的总页数
            int slideCount = presentationPart.SlideParts.Count();
            Console.WriteLine(slideCount.ToString());

            //获取幻灯片演示文稿子元素
            OpenXmlElementList slideList = presentationPart.Presentation.SlideIdList.ChildElements;
            for (int i = 0; i < slideCount; i++)
            {
                //判断要修改的幻灯片页码
                switch (i)
                {
                    case 0:
                        UpdateSlideNum1(presentationPart, i, slideList);
                        break;
                }
            }

            //关闭对ppt的修改， 如果没有关闭，修改后的ppt打不开
            presentationDocument.Close();
            presentationDocument.Dispose();
            MessageBox.Show("Test");
            //ProcessStartInfo processStartInfo = new ProcessStartInfo();
            //processStartInfo.Arguments = newFilePath;
            //processStartInfo.FileName = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk";
            //processStartInfo.WindowStyle = ProcessWindowStyle.Normal;
            //Process process = Process.Start(processStartInfo);
            //process.WaitForExit();
            //System.Environment.Exit(0);
        }

        static void UpdateSlideNum1(PresentationPart presentationPart,int slideIndex,OpenXmlElementList slideList)
        {
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            keyValues.Add("标题", "标题");
            keyValues.Add("副标题", "副标题");
            Dictionary<string, int> keyValuesWordSize = new Dictionary<string, int>();
            keyValuesWordSize.Add("标题", 9600);
            keyValuesWordSize.Add("副标题", 5000);
            Dictionary<string, string> keyValuesWordColor = new Dictionary<string, string>();
            keyValuesWordColor.Add("标题", "6e9020");
            UpdatePptWords(presentationPart, slideIndex, slideList,keyValues,wordFont: "Microsoft JhengHei",wordSize:keyValuesWordSize,fontColor:keyValuesWordColor);
        }

        /// <summary>
        /// 更新幻灯片中的字符
        /// </summary>
        /// <param name="presentationPart"></param>
        /// <param name="slideIndex">幻灯片页码</param>
        /// <param name="slideList">幻灯片集合</param>
        /// <param name="words">文字修改集合 key:对象name value：对象text</param>
        /// <param name="wordSize">字体大小,这里是全局的,这个幻灯片中所有文字大小都设置了</param>
        /// <param name="wordFont">字体</param>
        /// <param name="fontColor">字体颜色</param>
        static void UpdatePptWords(PresentationPart presentationPart, int slideIndex, OpenXmlElementList slideList, Dictionary<string, string> words,
            Dictionary<string, int> wordSize = null, string wordFont = null, Dictionary<string, string> fontColor = null)
        {
            //获取当前幻灯片的关系id
            string relationshipId = (slideList[slideIndex] as P.SlideId).RelationshipId;

            //根据id 获取当前幻灯片xml内容
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);

            //从Open Xml Sdk 中查看该文字在哪个xml节点，循环遍历该节点下面所有子节点
            P.ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            foreach (P.Shape shape in shapeTree.Descendants<P.Shape>())
            {
                //遍历
                foreach (var item in words)
                {
                    //从Open Xml Sdk 可以看到ppt中所设置的key在NonVisualDrawingProperties节点中
                    if (shape.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault().Name.Value ==item.Key)
                    {
                        //通过key可以获取该文字在D.Text节点中
                        shape.Descendants<D.Text>().FirstOrDefault().Text = item.Value;
                        if (wordFont != null&& shape.Descendants<D.LatinFont>().FirstOrDefault()!=null)
                        {
                            //同样是通过key可以在该shapetree中找到
                            shape.Descendants<D.LatinFont>().FirstOrDefault().Typeface = wordFont;
                        }
                        if (wordSize != null)
                        {
                            foreach (var wordSizeItem in wordSize)
                            {
                                //获取当前文字的size对象
                                if (shape.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault().Name.Value == wordSizeItem.Key)
                                    shape.Descendants<D.RunProperties>().FirstOrDefault().FontSize = wordSizeItem.Value;
                            }
                        }
                        if (fontColor != null)
                        {
                            foreach (var fontcolorItem in fontColor)
                            {
                                //获取当前文字的color
                                if (shape.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault().Name.Value == fontcolorItem.Key)
                                    shape.Descendants<D.SolidFill>().FirstOrDefault().RgbColorModelHex.Val = fontcolorItem.Value;
                            }
                        }
                    }
                }
              
            }
        }
    }
}
