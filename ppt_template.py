from PIL import Image
import datetime
import os

class VBACodeGenerator:

    def __init__(self):
        self.slide_width = 8.5*72
        self.slide_height = 18.0*72
        self.figure_width = 0.96*self.slide_width
        self.current_page = 1

    def crop_to_ratio(self, image_path, target_ratio):
        # Open the image using Pillow
        image = Image.open(image_path)

        # Get the width and height of the original image
        width, height = image.size

        # Calculate the current aspect ratio
        current_ratio = width / height

        if current_ratio > target_ratio:
            # Crop the width of the image
            new_width = height * target_ratio
            left = (width - new_width) / 2
            right = left + new_width
            top = 0
            bottom = height
        else:
            # Crop the height of the image
            new_height = width / target_ratio
            left = 0
            right = width
            top = (height - new_height) / 2
            bottom = top + new_height

        # Crop the image using the calculated coordinates
        cropped_image = image.crop((left, top, right, bottom))

        # Save the cropped image
        # cropped_image_path = os.path.splitext(image_path)[0] + "_cropped.jpg"
        # cropped_image_path = "./temp/image" + "_cropped.jpg"
        cropped_image_path = os.getcwd()+'/temp/image/'+"image_cropped.jpg"
        cropped_image.save(cropped_image_path)

        return cropped_image_path

    def concat_code(self, content_dict):
        head_str = """
        Sub CreatePowerPoint()
            Dim PowerPointApp As Object
            Dim Presentation As Object
            Dim Slide As Object
            Dim Shape As Object

            ' 创建PowerPoint应用程序对象
            Set PowerPointApp = CreateObject("PowerPoint.Application")

            ' 创建演示文稿
            Set Presentation = PowerPointApp.Presentations.Add
        """

        tail_str = """
            ' 显示PowerPoint窗口
            PowerPointApp.Visible = True

            ' 清理对象
            Set Shape = Nothing
            Set Slide = Nothing
            Set Presentation = Nothing
            Set PowerPointApp = Nothing
        End Sub
        """

        # set slide settings
        slide_setting = f"""
            ' 设置页面尺寸为8.5*18英寸
            Presentation.PageSetup.SlideWidth = {self.slide_width}
            Presentation.PageSetup.SlideHeight = {self.slide_height}

            ' 设置演示文稿的主题颜色
            Presentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(38, 38, 38)
        """
        cover_page = self.make_cover()

        # iterate through all pages
        content_pages = []
        for page_idx, page_content in content_dict.items():
            content_page = self.make_content(page_content)
            content_pages.append(content_page)
        content_str = "".join(content_pages)
        ending_page = self.make_end()
        macro = "".join([head_str, slide_setting, cover_page, content_str, ending_page, tail_str])
        return macro

    def make_cover(self):
        # read current path, check the latest episode number
        current_episode = 1
        # if os.path.exists(os.getcwd()):
        #     for file in os.listdir(os.getcwd()):
        #         if file.endswith('.pptx'):


        # image_path = r"C:\Users\htzha\Pictures\Screen Shot\夜撫でるメノウ.jpg"
        image_path = os.getcwd() + "/resource/amaranth.png"
        text_box_width = 600 / 612 * self.slide_width
        text_box_height = 50

        cover_page = f"""
            ' 新建页面
            Set Slide = Presentation.Slides.Add({self.current_page},11)

            ' 插入标题文本框
            Set Shape = Slide.Shapes.AddTextbox(1, {self.slide_width / 2 - text_box_width / 2}, {self.slide_height * 0.5}, {text_box_width}, {text_box_height})
            Shape.TextFrame.TextRange.Text = "长弓大盘鸡"
            Shape.TextFrame.TextRange.Font.Size = 96
            Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

            ' 插入副文本框
            Set Shape = Slide.Shapes.AddTextbox(1, {self.slide_width / 2 - text_box_width / 2}, {self.slide_height * 0.6}, {text_box_width}, {text_box_height})
            Shape.TextFrame.TextRange.Text = "第 {current_episode} 期"
            Shape.TextFrame.TextRange.Font.Size = 60
            Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

            ' 插入头像
            Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 0, 0)
            ' 锁定图片宽高比
            Shape.LockAspectRatio = msoTrue
            Shape.Width = 5 * 72
            slideWidth = Slide.Master.Width
            slideHeight = Slide.Master.Height
            Shape.Left = (slideWidth - Shape.Width) / 2
            Shape.Top = slideHeight * 0.3
        """

        self.current_page += 1
        return cover_page

    def make_content(self, input_dict):
        today = datetime.datetime.today()
        chapter_name = input_dict["Chapter_Name"]
        image_path = input_dict["Image_Path"]
        if image_path:
            image_path = self.crop_to_ratio(image_path, target_ratio=4 / 3)
        content_text = input_dict["Content_Text"]
        content_text = content_text.replace("\n", '"& vbNewLine &"')

        week_dir = {
            0: "一",
            1: "二",
            2: "三",
            3: "四",
            4: "五",
            5: "六",
            6: "日"
        }
        week = week_dir.get(today.weekday())
        today = today.strftime("%Y年%m月%d日")

        content_page = []
        # create new page
        content_page.append(
            f"""
            ' 新建页面
            Set Slide = Presentation.Slides.Add({self.current_page}, 11)
            
            """
        )

        # insert image if exist
        if image_path:
            content_page.append(
                f"""
                ' 插入背景图片
                Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 0, 0)
                ' 锁定图片宽高比
                Shape.LockAspectRatio = msoTrue
                Shape.Height = {self.slide_height}
                Shape.Fill.PictureEffects.Insert(msoEffectBlur).EffectParameters(1).Value = 100
                Shape.PictureFormat.Brightness = 0.2
                Shape.PictureFormat.Contrast = 0.2
                Shape.Left = {self.slide_width / 2} - Shape.Width/2
                
                ' 插入图片
                Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 100, 100)
                Shape.AutoShapeType = msoShapeRoundedRectangle
                Shape.Adjustments.Item(1) = 0.02 ' 替换 0.5 为你想要的圆角半径值
                Shape.LockAspectRatio = msoTrue ' 锁定图片宽高比
                Shape.Width = 592
                Shape.Left = {self.slide_width / 2} - Shape.Width/2
                Shape.Top = {self.slide_height * 0.235}
                
                """
            )

        # insert date
        content_page.append(
            f"""
            ' 插入日期文本框
            Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 0)
            Shape.TextFrame.TextRange.Text = "{today}"& vbNewLine &"星期{week}"
            Shape.TextFrame.TextRange.Font.Size = 36
            Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
            Shape.Left = {self.slide_width / 2 - self.figure_width / 2}
            Shape.Top = {self.slide_height * 0.15}
            
            """
        )

        # insert title if exist
        if chapter_name:
            content_page.append(
                f"""
                ' 插入对角圆角矩形（底）
                Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 10 , 100, 250, 80)
                Shape.Fill.ForeColor.RGB = RGB(174, 66, 7)
                Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
                Shape.Line.Visible = msoFalse
                Shape.left = {self.slide_width / 2 - self.figure_width / 2 + 0.01 * self.slide_width}
                Shape.Top = {self.slide_height * 0.15 - 0.01 * self.slide_width}
    
                ' 插入章节标题对角圆角矩形（表）
                Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 5 , 90, 250, 80)
                Shape.Fill.ForeColor.RGB = RGB(255, 165, 0)
                Shape.TextFrame.TextRange.Text = "{chapter_name}"
                Shape.TextFrame.TextRange.Font.Size = 42
                Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
                Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
                Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
                Shape.Line.Visible = msoFalse
                Shape.left = {self.slide_width / 2 - self.figure_width / 2}
                Shape.Top = {self.slide_height * 0.15}
                
                """
            )
        else:
            content_page.append(
                f"""
                ' 插入对角圆角矩形（底）
                Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 10 , 100, 250, 80)
                Shape.Fill.ForeColor.RGB = RGB(174, 66, 7)
                Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
                Shape.Line.Visible = msoFalse
                Shape.left = {self.slide_width / 2 - self.figure_width / 2 + 0.01 * self.slide_width}
                Shape.Top = {self.slide_height * 0.15 - 0.01 * self.slide_width}

                ' 插入章节标题对角圆角矩形（表）
                Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 5 , 90, 250, 80)
                Shape.Fill.ForeColor.RGB = RGB(255, 165, 0)
                Shape.TextFrame.TextRange.Font.Size = 42
                Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
                Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
                Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
                Shape.Line.Visible = msoFalse
                Shape.left = {self.slide_width / 2 - self.figure_width / 2}
                Shape.Top = {self.slide_height * 0.15}
                
                """
            )

        # insert separation bar
        content_page.append(
            f"""
            ' 插入圆角矩形（分隔栏）
            Set Shape = Slide.Shapes.AddShape(msoShapeRoundedRectangle, 10 , 100, {self.figure_width}, 2)
            Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Shape.Fill.Transparency = 0.5
            Shape.Line.Visible = msoFalse
            Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
            Shape.left = {self.slide_width / 2 - self.figure_width / 2}
            Shape.Top = {self.slide_height * 0.59}
            
            """
        )

        # insert text if exist
        if content_text:
            content_page.append(
                f"""
                ' 插入内容文本框
                Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 200)
                Shape.TextFrame.TextRange.Text = "{content_text}"
                Shape.TextFrame.TextRange.Font.Size = 33
                Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Medium"
                Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0"
                Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                Shape.Left = {self.slide_width / 2} - Shape.Width/2
                Shape.Top = {self.slide_height * 0.6}
                Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignJustify
                
                """
            )

        content_page = "".join(content_page)


        # content_page = f"""
        #     ' 新建页面
        #     Set Slide = Presentation.Slides.Add({self.current_page}, 11)
        #
        #     ' 插入背景图片
        #     Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 0, 0)
        #     ' 锁定图片宽高比
        #     Shape.LockAspectRatio = msoTrue
        #     Shape.Height = {self.slide_height}
        #     Shape.Fill.PictureEffects.Insert(msoEffectBlur).EffectParameters(1).Value = 100
        #     Shape.PictureFormat.Brightness = 0.2
        #     Shape.PictureFormat.Contrast = 0.2
        #     Shape.Left = {self.slide_width / 2} - Shape.Width/2
        #
        #     ' 插入日期文本框
        #     Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 0)
        #     Shape.TextFrame.TextRange.Text = "{today}"& vbNewLine &"星期{week}"
        #     Shape.TextFrame.TextRange.Font.Size = 36
        #     Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
        #     Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
        #     Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        #     Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
        #     Shape.Left = {self.slide_width / 2 - self.figure_width / 2}
        #     Shape.Top = {self.slide_height * 0.15}
        #
        #     ' 插入对角圆角矩形（底）
        #     Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 10 , 100, 250, 80)
        #     Shape.Fill.ForeColor.RGB = RGB(174, 66, 7)
        #     Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
        #     Shape.Line.Visible = msoFalse
        #     Shape.left = {self.slide_width / 2 - self.figure_width / 2 + 0.01 * self.slide_width}
        #     Shape.Top = {self.slide_height * 0.15 - 0.01 * self.slide_width}
        #
        #     ' 插入章节标题对角圆角矩形（表）
        #     Set Shape = Slide.Shapes.AddShape(msoShapeRound2DiagRectangle, 5 , 90, 250, 80)
        #     Shape.Fill.ForeColor.RGB = RGB(255, 165, 0)
        #     Shape.TextFrame.TextRange.Text = "{chapter_name}"
        #     Shape.TextFrame.TextRange.Font.Size = 42
        #     Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
        #     Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
        #     Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        #     Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
        #     Shape.Line.Visible = msoFalse
        #     Shape.left = {self.slide_width / 2 - self.figure_width / 2}
        #     Shape.Top = {self.slide_height * 0.15}
        #
        #     ' 插入圆角矩形（分隔栏）
        #     Set Shape = Slide.Shapes.AddShape(msoShapeRoundedRectangle, 10 , 100, {self.figure_width}, 2)
        #     Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        #     Shape.Fill.Transparency = 0.5
        #     Shape.Line.Visible = msoFalse
        #     Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
        #     Shape.left = {self.slide_width / 2 - self.figure_width / 2}
        #     Shape.Top = {self.slide_height * 0.59}
        #
        #     ' 插入图片
        #     Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 100, 100)
        #     Shape.AutoShapeType = msoShapeRoundedRectangle
        #     Shape.Adjustments.Item(1) = 0.02 ' 替换 0.5 为你想要的圆角半径值
        #     Shape.LockAspectRatio = msoTrue ' 锁定图片宽高比
        #     Shape.Width = 592
        #     Shape.Left = {self.slide_width / 2} - Shape.Width/2
        #     Shape.Top = {self.slide_height * 0.235}
        #
        #     ' 插入内容文本框
        #     Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 200)
        #     Shape.TextFrame.TextRange.Text = "{content_text}"
        #     Shape.TextFrame.TextRange.Font.Size = 33
        #     Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Medium"
        #     Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0"
        #     Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        #     Shape.Left = {self.slide_width / 2} - Shape.Width/2
        #     Shape.Top = {self.slide_height * 0.6}
        #     Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignJustify
        # """

        self.current_page += 1
        return content_page

    def make_end(self):
        image_path = r"C:\Users\htzha\Pictures\Screen Shot\warning.svg"
        warning_text = "视频仅作个人经验分享，不构成任何投资建议，对于跟随本视频投资造成的损失概不负责。"
        ending_page = f"""
            ' 新建页面
            Set Slide = Presentation.Slides.Add({self.current_page}, 11)

            ' 插入图片
            Set Shape = Slide.Shapes.AddPicture("{image_path}", msoFalse, msoTrue, 10, 100)
            Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Shape.Fill.Transparency = 0.3
            Shape.LockAspectRatio = msoTrue ' 锁定图片宽高比
            Shape.Width = 150
            Shape.Left = {self.slide_width / 2} - Shape.Width/2
            Shape.Top = {self.slide_height * 0.25}

            ' 插入圆角矩形（分隔栏）
            Set Shape = Slide.Shapes.AddShape(msoShapeRoundedRectangle, 10 , 100, {self.figure_width}, 2)
            Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Shape.Fill.Transparency = 0.5
            Shape.Line.Visible = msoFalse
            Shape.Adjustments.Item(1) = 1 ' 替换 0.5 为你想要的圆角半径值
            Shape.Top = {self.slide_height * 0.4}

            ' 插入内容文本框
            Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 200)
            Shape.TextFrame.TextRange.Text = "{warning_text}"
            Shape.TextFrame.TextRange.Font.Size = 33
            Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Medium"
            Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0"
            Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            With Shape.TextFrame.TextRange.ParagraphFormat
                .LineRuleWithin = msoTrue
                .SpaceWithin = 1.5    '行间距
            End With
            Shape.Left = {self.slide_width / 2} - Shape.Width/2
            Shape.Top = {self.slide_height * 0.4}
            Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignJustify

            ' 插入内容文本框
            Set Shape = Slide.Shapes.AddTextbox(1, 100, 100, {self.figure_width}, 200)
            Shape.TextFrame.TextRange.Text = "股市有风险 投资需谨慎"
            Shape.TextFrame.TextRange.Font.Size = 50
            Shape.TextFrame.TextRange.Font.NameFarEast = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Name = "OPlusSans 3.0 Bold"
            Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Shape.Left = {self.slide_width / 2} - Shape.Width/2
            Shape.Top = {self.slide_height * 0.6}
            Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        """

        self.current_page += 1
        return ending_page
