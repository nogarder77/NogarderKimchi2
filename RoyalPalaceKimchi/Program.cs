using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace RoyalPalaceKimchi
{
    class Program
    {
        #region 상수 정의
        private const string Version = "1.0.0";
        private const string Author = "김용식";
        private const string Email = "nogarder@gmail.com";
        private const string Copyright = "© 2025 김용식, All Rights Reserved";
        
        // 명령어 상수
        private const string CMD_SAVE = "-save";
        private const string CMD_PAGE = "-page";
        private const string CMD_PRINT = "-print";
        private const string CMD_HELP = "-help";
        private const string CMD_HELP_ALT = "/?";
        private const string CMD_VERSION = "-version";
        
        // 검색 연산자
        private const char OP_AND = '&';
        private const char OP_OR = '|';
        #endregion

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            if (args.Length == 0 || IsHelpCommand(args[0]))
            {
                ShowHelp();
                return;
            }

            if (IsVersionCommand(args[0]))
            {
                ShowVersion();
                return;
            }

            if (args.Length < 3)
            {
                Console.WriteLine("오류: 인수가 충분하지 않습니다.");
                Console.WriteLine("사용법: RoyalPalaceKimchi.exe [PDF파일경로] [검색문자열] [명령]");
                Console.WriteLine($"자세한 정보는 {CMD_HELP} 또는 {CMD_HELP_ALT} 명령어를 사용하세요.");
                return;
            }

            string pdfFilePath = args[0];
            string searchQuery = args[1];
            string command = args[2].ToLower();

            try
            {
                if (!File.Exists(pdfFilePath))
                {
                    Console.WriteLine($"오류: '{pdfFilePath}' 파일을 찾을 수 없습니다.");
                    return;
                }

                List<int> matchedPages = SearchPdf(pdfFilePath, searchQuery);

                if (matchedPages.Count == 0)
                {
                    Console.WriteLine("검색 결과: 일치하는 페이지가 없습니다.");
                    return;
                }

                Console.WriteLine($"검색 결과: {matchedPages.Count}개의 페이지에서 일치하는 내용이 발견되었습니다.");

                switch (command)
                {
                    case CMD_SAVE:
                        SaveFilteredPdf(pdfFilePath, matchedPages, searchQuery);
                        break;
                    case CMD_PAGE:
                        ShowMatchedPages(matchedPages);
                        break;
                    case CMD_PRINT:
                        PrintMatchedPages(pdfFilePath, matchedPages);
                        break;
                    default:
                        Console.WriteLine($"오류: 알 수 없는 명령어 '{command}'");
                        Console.WriteLine($"지원되는 명령어: {CMD_SAVE}, {CMD_PAGE}, {CMD_PRINT}");
                        Console.WriteLine($"자세한 정보는 {CMD_HELP} 또는 {CMD_HELP_ALT} 명령어를 사용하세요.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류가 발생했습니다: {ex.Message}");
                Console.WriteLine("자세한 정보는 로그를 확인하세요.");
            }
        }

        #region 명령어 및 도움말
        private static bool IsHelpCommand(string arg)
        {
            return arg.Equals(CMD_HELP, StringComparison.OrdinalIgnoreCase) || 
                   arg.Equals(CMD_HELP_ALT, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsVersionCommand(string arg)
        {
            return arg.Equals(CMD_VERSION, StringComparison.OrdinalIgnoreCase);
        }

        private static void ShowHelp()
        {
            Console.WriteLine("RoyalPalaceKimchi - PDF 검색 및 필터링 도구");
            Console.WriteLine($"버전: {Version}");
            Console.WriteLine($"{Copyright}");
            Console.WriteLine();
            Console.WriteLine("사용법: RoyalPalaceKimchi.exe [PDF파일경로] [검색문자열] [명령]");
            Console.WriteLine();
            Console.WriteLine("명령어:");
            Console.WriteLine($"  {CMD_SAVE}    검색된 페이지만 포함하는 새 PDF 파일 생성");
            Console.WriteLine($"  {CMD_PAGE}    검색된 페이지 번호 목록을 콘솔에 출력");
            Console.WriteLine($"  {CMD_PRINT}   검색된 페이지를 기본 프린터로 인쇄");
            Console.WriteLine($"  {CMD_HELP}    이 도움말 메시지 표시");
            Console.WriteLine($"  {CMD_HELP_ALT}       이 도움말 메시지 표시");
            Console.WriteLine($"  {CMD_VERSION} 버전 정보 표시");
            Console.WriteLine();
            Console.WriteLine("검색 연산자:");
            Console.WriteLine($"  {OP_AND}        AND 연산자 (예: \"단어1 {OP_AND} 단어2\")");
            Console.WriteLine($"  {OP_OR}        OR 연산자 (예: \"단어1 {OP_OR} 단어2\")");
            Console.WriteLine();
            Console.WriteLine("예시:");
            Console.WriteLine($"  RoyalPalaceKimchi.exe \"C:\\문서\\보고서.pdf\" \"예산 {OP_AND} 2024\" {CMD_SAVE}");
            Console.WriteLine($"  RoyalPalaceKimchi.exe \"보고서.pdf\" \"수출 {OP_OR} 수입\" {CMD_PAGE}");
        }

        private static void ShowVersion()
        {
            Console.WriteLine($"RoyalPalaceKimchi 버전 {Version}");
            Console.WriteLine($"제작자: {Author} ({Email})");
            Console.WriteLine(Copyright);
        }
        #endregion

        #region PDF 검색 및 처리
        /// <summary>
        /// PDF 파일에서 검색어가 포함된 페이지를 찾습니다.
        /// </summary>
        private static List<int> SearchPdf(string pdfFilePath, string searchQuery)
        {
            List<int> matchedPages = new List<int>();
            searchQuery = PrepareSearchQuery(searchQuery);
            
            try
            {
                // PDF 파일 열기 및 처리
                using (PdfReader reader = OpenPdfReader(pdfFilePath))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    int totalPages = pdfDoc.GetNumberOfPages();
                    Console.WriteLine($"PDF 파일 내 총 {totalPages}페이지를 검색 중...");
                    
                    for (int i = 1; i <= totalPages; i++)
                    {
                        try
                        {
                            var page = pdfDoc.GetPage(i);
                            // 텍스트 추출 전략 변경 - LocationTextExtractionStrategy는 위치 기반 텍스트 추출
                            string text = PdfTextExtractor.GetTextFromPage(page, new LocationTextExtractionStrategy());
                            
                            if (EvaluateSearchExpression(text, searchQuery))
                            {
                                matchedPages.Add(i);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"페이지 {i} 처리 중 오류: {ex.Message}");
                            // 오류 발생해도 계속 진행
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF 처리 중 오류 발생: {ex.Message}");
                throw;
            }
            
            return matchedPages;
        }

        /// <summary>
        /// 검색어를 소문자로 변환하여 대소문자 구분 없이 검색할 수 있도록 합니다.
        /// </summary>
        private static string PrepareSearchQuery(string searchQuery)
        {
            // 검색어를 소문자로 변환 (대소문자 구분 없음)
            return searchQuery.ToLower();
        }

        /// <summary>
        /// 논리 연산자를 포함한 검색식을 평가합니다.
        /// </summary>
        private static bool EvaluateSearchExpression(string text, string searchExpression)
        {
            text = text.ToLower(); // 대소문자 구분 없이 검색하기 위해
            
            // OR 연산 처리
            if (searchExpression.Contains(OP_OR))
            {
                string[] orTerms = searchExpression.Split(OP_OR);
                foreach (string term in orTerms)
                {
                    if (EvaluateSearchExpression(text, term.Trim()))
                        return true;
                }
                return false;
            }
            
            // AND 연산 처리
            if (searchExpression.Contains(OP_AND))
            {
                string[] andTerms = searchExpression.Split(OP_AND);
                foreach (string term in andTerms)
                {
                    if (!EvaluateSearchExpression(text, term.Trim()))
                        return false;
                }
                return true;
            }
            
            // 단일 검색어 처리
            return text.Contains(searchExpression);
        }
        #endregion

        #region 결과 처리 (저장, 출력, 인쇄)
        /// <summary>
        /// 선택된 페이지만 포함하는 새 PDF 파일을 저장합니다.
        /// </summary>
        private static void SaveFilteredPdf(string originalPdfPath, List<int> pagesToInclude, string searchQuery)
        {
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(originalPdfPath);
                string newFileName = $"{fileName}_{CleanSearchQueryForFileName(searchQuery)}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.pdf";
                string outputPath = Path.Combine(Path.GetDirectoryName(originalPdfPath) ?? "", newFileName);

                CopyPdfPages(originalPdfPath, outputPath, pagesToInclude);
                
                Console.WriteLine($"필터링된 PDF 파일이 저장되었습니다: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF 저장 중 오류 발생: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"원인: {ex.InnerException.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// 파일명으로 사용할 수 없는 문자를 제거합니다.
        /// </summary>
        private static string CleanSearchQueryForFileName(string searchQuery)
        {
            // 파일명에 사용할 수 없는 문자 제거
            return Regex.Replace(searchQuery, @"[\\/:*?""<>|&]", "_");
        }

        /// <summary>
        /// 검색된 페이지 번호 목록을 출력합니다.
        /// </summary>
        private static void ShowMatchedPages(List<int> matchedPages)
        {
            Console.WriteLine("검색된 페이지:");
            Console.WriteLine(string.Join(", ", matchedPages));
        }

        /// <summary>
        /// 선택된 페이지만 인쇄합니다.
        /// </summary>
        private static void PrintMatchedPages(string pdfFilePath, List<int> pagesToPrint)
        {
            try
            {
                Console.WriteLine("Windows 기본 프린터로 인쇄 중...");
                
                // 프린터 존재 여부 확인
                if (!IsPrinterAvailable())
                {
                    Console.WriteLine("경고: 시스템에 기본 프린터가 설정되어 있지 않습니다.");
                    Console.WriteLine("인쇄를 계속 진행하시겠습니까? (Y/N)");
                    
                    var key = Console.ReadKey(true);
                    if (key.Key != ConsoleKey.Y)
                    {
                        Console.WriteLine("인쇄가 취소되었습니다.");
                        return;
                    }
                    
                    Console.WriteLine("인쇄를 계속 진행합니다...");
                }
                
                // 임시 PDF 파일 생성
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"RoyalPalaceKimchi_Temp_{Guid.NewGuid()}.pdf");
                Console.WriteLine($"임시 PDF 파일 생성: {tempPdfPath}");
                
                // 필터링된 페이지만 포함하는 임시 PDF 저장
                CopyPdfPages(pdfFilePath, tempPdfPath, pagesToPrint);
                
                // 시스템 프린팅 서비스 사용하여 인쇄
                if (File.Exists(tempPdfPath))
                {
                    bool printResult = PrintDocument(tempPdfPath);
                    
                    // 임시 파일 삭제
                    try 
                    { 
                        // 인쇄 작업 중일 때 파일이 잠길 수 있으므로 잠시 대기
                        System.Threading.Thread.Sleep(1000);
                        File.Delete(tempPdfPath); 
                    } 
                    catch (Exception ex) 
                    { 
                        Console.WriteLine($"임시 파일 삭제 중 오류: {ex.Message}"); 
                    }
                    
                    if (printResult)
                    {
                        Console.WriteLine("인쇄가 성공적으로 요청되었습니다.");
                    }
                    else
                    {
                        Console.WriteLine("인쇄 요청 중 문제가 발생했습니다.");
                    }
                }
                else
                {
                    Console.WriteLine("오류: 임시 PDF 파일이 생성되지 않았습니다.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"인쇄 중 오류가 발생했습니다: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"원인: {ex.InnerException.Message}");
                }
            }
        }
        #endregion
        
        #region 유틸리티 메서드
        /// <summary>
        /// 한글 PDF 지원을 위한 설정을 포함한 PdfReader를 생성합니다.
        /// </summary>
        private static PdfReader OpenPdfReader(string filePath)
        {
            var reader = new PdfReader(filePath);
            reader.SetUnethicalReading(true); // 한글 PDF 지원을 위한 설정
            return reader;
        }
        
        /// <summary>
        /// 원본 PDF에서 선택된 페이지만 새 PDF로 복사합니다.
        /// </summary>
        private static void CopyPdfPages(string sourcePdfPath, string targetPdfPath, List<int> pages)
        {
            Console.WriteLine($"페이지 복사 중: {string.Join(", ", pages)}");
            
            using (PdfReader reader = OpenPdfReader(sourcePdfPath))
            using (PdfDocument inputDoc = new PdfDocument(reader))
            using (PdfWriter writer = new PdfWriter(targetPdfPath))
            using (PdfDocument outputDoc = new PdfDocument(writer))
            {
                foreach (int pageNum in pages)
                {
                    try
                    {
                        // 페이지 복사
                        inputDoc.CopyPagesTo(pageNum, pageNum, outputDoc);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"페이지 {pageNum} 복사 중 오류: {ex.Message}");
                    }
                }
            }
        }
        
        /// <summary>
        /// 시스템에 프린터가 설정되어 있는지 확인합니다.
        /// </summary>
        private static bool IsPrinterAvailable()
        {
            try
            {
                // Windows 플랫폼에서만 프린터 설정 확인
                if (OperatingSystem.IsWindows())
                {
                    // Windows 6.1 이상에서만 지원됨 (Windows 7/Server 2008 R2 이상)
                    if (OperatingSystem.IsWindowsVersionAtLeast(6, 1))
                    {
                        return CheckPrinters();
                    }
                    else
                    {
                        Console.WriteLine("경고: 프린터 확인은 Windows 7 이상에서만 지원됩니다.");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine("경고: 프린터 확인은 Windows 운영체제에서만 지원됩니다.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"프린터 확인 중 오류: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Windows 6.1 이상에서 사용 가능한 프린터를 확인합니다.
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows6.1")]
        private static bool CheckPrinters()
        {
            System.Drawing.Printing.PrinterSettings.StringCollection? printers = 
                System.Drawing.Printing.PrinterSettings.InstalledPrinters;
            
            if (printers != null && printers.Count > 0)
            {
                var defaultPrinter = printers.Cast<string>().FirstOrDefault();
                if (!string.IsNullOrEmpty(defaultPrinter))
                {
                    Console.WriteLine($"기본 프린터: {defaultPrinter}");
                    return true;
                }
            }
            
            return false;
        }
        
        /// <summary>
        /// PDF 파일을 기본 프린터로 인쇄합니다.
        /// </summary>
        private static bool PrintDocument(string filePath)
        {
            try
            {
                // Windows 기본 프린터로 인쇄
                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    CreateNoWindow = true,
                    UseShellExecute = true
                };
                
                using (var process = System.Diagnostics.Process.Start(processStartInfo))
                {
                    if (process == null)
                    {
                        Console.WriteLine("인쇄 프로세스를 시작할 수 없습니다.");
                        return false;
                    }
                    
                    // 프로세스가 시작되었는지 확인
                    if (!process.HasExited)
                    {
                        // 최대 5초 동안 대기 (기본 종료 대기 시간 증가)
                        if (!process.WaitForExit(5000))
                        {
                            Console.WriteLine("인쇄 프로세스가 예상보다 오래 실행 중입니다. 백그라운드에서 계속 진행됩니다.");
                        }
                    }
                    
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"인쇄 문서 처리 중 오류: {ex.Message}");
                return false;
            }
        }
        #endregion
    }
}
