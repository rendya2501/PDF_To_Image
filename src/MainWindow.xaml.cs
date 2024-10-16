using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using PdfiumViewer;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;

namespace PDF_to_Image;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// PDF選択ダイアログを表示
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OpenPdfFileDialogButton_Click(object sender, RoutedEventArgs e)
    {
        // OpenFileDialog インスタンスを作成
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            // PDFファイルのみ選択できるようにフィルタを設定
            Filter = "PDF files (*.pdf)|*.pdf",
            // 複数のファイルを選択できないように設定
            Multiselect = false
        };

        // ダイアログを表示
        if (openFileDialog.ShowDialog() != true) return;

        // 選択したファイルのパスを取得
        string selectedFilePath = openFileDialog.FileName;
        this.Path.Text = selectedFilePath;
        this.DirectoryPath.Text = @$"{selectedFilePath[..selectedFilePath.LastIndexOf('\\')]}\output";
    }

    /// <summary>
    /// 画像出力ディレクトリダイアログを表示
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OpenDirectryDialogButton_Click(object sender, RoutedEventArgs e)
    {
        using CommonOpenFileDialog cofd = new CommonOpenFileDialog();
        // フォルダを選択できるようにする
        cofd.IsFolderPicker = true;

        if (cofd.ShowDialog() == CommonFileDialogResult.Ok)
        {
            this.DirectoryPath.Text = cofd.FileName;
        }
    }

    /// <summary>
    /// pdfを画像に変換する処理を実行
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void ExecutePDFToImage_Click(object sender, RoutedEventArgs e)
    {
        if (!Validation()) return;

        // クルクルを表示
        this.LoadingOverlay.Visibility = Visibility.Visible;
        this.Whole.IsEnabled = false;

        // pdfのファイルパス
        string inputPath = this.Path.Text;
        // 結合後の画像の保存フォルダ
        string outputFolderPath = this.DirectoryPath.Text;

        // 出力フォルダが存在しない場合は作成
        if (!Directory.Exists(outputFolderPath))
        {
            Directory.CreateDirectory(outputFolderPath);
        }

        // チェックされていなければ1ページモード
        // チェックされていれば2ページモード
        IProcessStrategy strategy = this.IsPairCheckBox.IsChecked != true
            // 1ページモード
            ? new ProcessStrategyA()
            // 2ページモード
            : new ProcessStrategyB();
        // 戦略の決定
        CheckBoxContext context = new CheckBoxContext();
        context.SetStrategy(strategy);

        // 実行
        await Task.Run(() => context.ExecuteStrategy(inputPath, outputFolderPath));

        // 処理完了後、クルクルを非表示
        this.LoadingOverlay.Visibility = Visibility.Collapsed;
        this.Whole.IsEnabled = true;
        MessageBox.Show("処理終了");
    }

    /// <summary>
    /// 実行前バリデーション
    /// </summary>
    /// <returns></returns>
    private bool Validation()
    {
        // エラーメッセージと対応する条件のディクショナリ
        var validations = new Dictionary<Func<bool>, string>
        {
            { () => string.IsNullOrEmpty(this.Path.Text), "PDFファイルが指定されていません。" },
            { () => !File.Exists(this.Path.Text), "存在しないPDFファイルです。" },
            { () => string.IsNullOrEmpty(this.DirectoryPath.Text), "出力ディレクトリが指定されていません。" }
        };

        // バリデーションチェック
        foreach (var validation in validations)
        {
            if (!validation.Key())
            {
                continue;
            }
            // バリデーションに引っかかったらメッセージを表示してfalseを返す
            MessageBox.Show(validation.Value);
            return false;
        }

        return true;
    }
}


// ストラテジーインターフェース
public interface IProcessStrategy
{
    void Execute(string inputPath, string outputFolderPath);
}

// 具体的なストラテジー1
public class ProcessStrategyA : IProcessStrategy
{
    public void Execute(string inputPath, string outputFolderPath)
    {
        using PdfDocument pdf = PdfDocument.Load(inputPath);
        //PDFの各ページをループする
        for (int i = 0; i < pdf.PageCount; i++)
        {
            //PDFページを画像に変換する。
            Image image = pdf.Render(i, 500, 500, PdfRenderFlags.CorrectFromDpi);

            //画像をjpeg形式で指定フォルダに保存 
            string file = @$"{outputFolderPath}\{i:D4}.jpg";
            image.Save(file, ImageFormat.Jpeg);
        }
    }
}

// 具体的なストラテジー2
public class ProcessStrategyB : IProcessStrategy
{
    public void Execute(string inputPath, string outputFolderPath)
    {
        using PdfDocument pdf = PdfDocument.Load(inputPath);

        int counter = 1;

        for (int i = 0; i < pdf.PageCount; i += 2)
        {
            using Image img1 = pdf.Render(i, 500, 500, PdfRenderFlags.CorrectFromDpi);
            using Image img2 = i + 1 < pdf.PageCount
                ? pdf.Render(i + 1, 500, 500, PdfRenderFlags.CorrectFromDpi)
                : new Func<Image>(() =>
                {
                    // 指定したサイズのBitmapを作成
                    Bitmap bitmap = new Bitmap(img1.Width, img1.Height);
                    // Graphicsオブジェクトを使ってBitmapを白で塗りつぶす
                    using Graphics g = Graphics.FromImage(bitmap);
                    g.Clear(Color.White);
                    // BitmapをImageとして返す
                    return bitmap;
                })();


            // 画像の結合
            int width = img1.Width + img2.Width;
            int height = img1.Height > img2.Height ? img1.Height : img2.Height;

            int newHeight1 = img1.Height * height / img1.Height;
            int newWidth1 = img1.Width * height / img1.Height;
            Rectangle rect1 = new Rectangle(0, 0, newWidth1, newHeight1);

            int newHeight2 = img2.Height * height / img2.Height;
            int newWidth2 = img2.Width * height / img2.Height;
            Rectangle rect2 = new Rectangle(newWidth1, 0, newWidth2, newHeight2);

            Bitmap bmp = new Bitmap(newWidth1 + newWidth2, height);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(img1, rect1);
            g.DrawImage(img2, rect2);
            g.Dispose();
            img1.Dispose();
            img2.Dispose();

            string file = @$"{outputFolderPath}\ToImage-{i:D4}.jpg";
            bmp.Save(file, ImageFormat.Jpeg);
            bmp.Dispose();

            counter++;
        }
    }
}

// Contextクラス
public class CheckBoxContext
{
    private IProcessStrategy _strategy;

    // チェックボックスの選択に応じてストラテジーをセット
    public void SetStrategy(IProcessStrategy strategy)
    {
        _strategy = strategy;
    }

    // ストラテジーに基づいて処理を実行
    public void ExecuteStrategy(string inputPath, string outputFolderPath)
    {
        _strategy?.Execute(inputPath, outputFolderPath);
    }
}
