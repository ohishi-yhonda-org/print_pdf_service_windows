package main

import (
	"bufio"
	"bytes"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime" // runtimeパッケージを追加
	"strconv"
	"strings"

	"github.com/getlantern/systray" // systrayライブラリを追加
)

// startHTTPServer はHTTPサーバーを起動する関数です。
func startHTTPServer() {
	log.Println("HTTPサーバーを開始します。")                    // ログ出力
	fmt.Println("Entering startHTTPServer function.") // デバッグ用: 関数開始をコンソールに出力

	// ルートパス ("/") へのリクエストを処理するハンドラを設定します。
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		fmt.Fprintf(w, "Hello from Go HTTP Server running as a Tray Application!")
	})

	// PDF印刷用の新しいハンドラを追加
	http.HandleFunc("/print-pdf", printPDFHandler)
	log.Println("/print-pdf ハンドラを追加しました。")   // ログ出力
	fmt.Println("Added /print-pdf handler.") // デバッグ用ログ

	// HTTPサーバーがリッスンするポートを設定します。
	port := ":8080"
	log.Printf("HTTPサーバーをポート %s で開始しようとしています。\n", port)              // ログ出力
	fmt.Printf("Attempting to start HTTP server on port %s\n", port) // デバッグ用: ポート情報をコンソールに出力

	// HTTPサーバーを起動し、エラーがあれば処理します。
	// ListenAndServeはブロッキング関数であり、エラーが発生した場合（ポートが使用中など）にのみ処理が返ります。
	if err := http.ListenAndServe(port, nil); err != nil {
		// log.Fatalfはプログラムを終了させるため、log.Printfに変更します。
		// これにより、サーバーの起動に失敗してもタスクトレイアプリは動作し続けます。
		log.Printf("HTTPサーバーの起動に失敗しました: %v", err)
		fmt.Printf("HTTP server failed: %v\n", err)
		// エラーが発生したことをユーザーに通知するためにツールチップを更新します。
		systray.SetTooltip(fmt.Sprintf("Go HTTP Printer (エラー: %v)", err))
	}
}

// printPDFHandler は /print-pdf エンドポイントのリクエストを処理します。
// POST multipart/form-data のボディからPDFファイルとプリンター名を受け取り、PDFを印刷します。
func printPDFHandler(w http.ResponseWriter, r *http.Request) {
	log.Println("Received request for /print-pdf.") // ログ出力
	fmt.Println("Received request for /print-pdf.") // デバッグ用ログ

	// POSTメソッド以外は許可しない
	if r.Method != http.MethodPost {
		http.Error(w, "このエンドポイントではPOSTメソッドのみ許可されています。", http.StatusMethodNotAllowed)
		log.Println("エラー: POSTメソッドのみ許可されています。")            // ログ出力
		fmt.Println("Error: Only POST method is allowed.") // デバッグ用ログ
		return
	}

	// multipart/form-data をパースします。最大10MBのファイルを受け入れます。
	err := r.ParseMultipartForm(10 << 20) // 10MB
	if err != nil {
		http.Error(w, fmt.Sprintf("マルチパートフォームのパースに失敗しました: %v", err), http.StatusBadRequest)
		log.Printf("エラー: マルチパートフォームのパースに失敗しました: %v\n", err)            // ログ出力
		fmt.Printf("Error: Failed to parse multipart form: %v\n", err) // デバッグ用ログ
		return
	}

	// プリンター名を取得します。
	printerName := r.FormValue("printer")
	if printerName == "" {
		http.Error(w, "フォームデータに'printer'パラメータがありません。", http.StatusBadRequest)
		log.Println("エラー: 'printer'パラメータがありません。")          // ログ出力
		fmt.Println("Error: Missing 'printer' parameter.") // デバッグ用ログ
		return
	}

	// アップロードされたPDFファイルを取得します。
	file, handler, err := r.FormFile("document")
	if err != nil {
		http.Error(w, fmt.Sprintf("フォームデータに'document'ファイルがありません: %v", err), http.StatusBadRequest)
		log.Printf("エラー: 'document'ファイルがありません: %v\n", err)      // ログ出力
		fmt.Printf("Error: Missing 'document' file: %v\n", err) // デバッグ用ログ
		return
	}
	defer file.Close() // 関数終了時にファイルを閉じます。

	// 一時ファイルを作成し、アップロードされたPDFを書き込みます。
	// tempDir := os.TempDir() // システムの一時ディレクトリを使用
	//現在のディレクトリに一時ファイルを作成します。
	// tempDir, err := os.Getwd() // 現在の作業ディレクトリを取得
	//root ディレクトリ(c:\)に一時ファイルを作成します。
	tempDir := "c:\\pdf"             // Windowsのルートディレクトリに一時ディレクトリを指定
	err = os.MkdirAll(tempDir, 0755) // ディレクトリが存在しない場合は作成します。{

	// := os.Mkdir("c:\\pdf", 0755) // c:\pdf ディレクトリを作成

	if err != nil {
		http.Error(w, fmt.Sprintf("一時ディレクトリの取得に失敗しました: %v", err), http.StatusInternalServerError)
		log.Printf("エラー: 一時ディレクトリの取得に失敗しました: %v\n", err)                  // ログ出力
		fmt.Printf("Error: Failed to get temporary directory: %v\n", err) // デバッグ用ログ
		return
	}
	tempFilePath := filepath.Join(tempDir, handler.Filename)
	log.Printf("アップロードされたファイルを一時パスに保存しています: %s\n", tempFilePath)             // ログ出力
	fmt.Printf("Saving uploaded file to temporary path: %s\n", tempFilePath) // デバッグ用ログ

	tempFile, err := os.Create(tempFilePath)
	if err != nil {
		http.Error(w, fmt.Sprintf("一時ファイルの作成に失敗しました: %v", err), http.StatusInternalServerError)
		log.Printf("エラー: 一時ファイルの作成に失敗しました: %v\n", err)                  // ログ出力
		fmt.Printf("Error: Failed to create temporary file: %v\n", err) // デバッグ用ログ
		return
	}

	_, err = io.Copy(tempFile, file)
	// io.Copy の後にファイルを明示的に閉じる必要があります。
	// これにより、Acrobat.exeがファイルにアクセスできるようになります。
	tempFile.Close()

	if err != nil {
		http.Error(w, fmt.Sprintf("アップロードされたファイルの保存に失敗しました: %v", err), http.StatusInternalServerError)
		log.Printf("エラー: アップロードされたファイルの保存に失敗しました: %v\n", err)        // ログ出力
		fmt.Printf("Error: Failed to save uploaded file: %v\n", err) // デバッグ用ログ
		return
	}
	log.Printf("一時ファイルに正常に保存しました: %s", tempFilePath) // ログ出力

	fmt.Printf("Attempting to print document '%s' to printer '%s'.\n", tempFilePath, printerName) // デバッグ用ログ

	// PDF印刷を実行します。
	err = printPDF(tempFilePath, printerName)
	if err != nil {
		http.Error(w, fmt.Sprintf("PDFの印刷に失敗しました: %v", err), http.StatusInternalServerError)
		log.Printf("PDF印刷エラー: %v\n", err)           // ログ出力
		fmt.Printf("Error printing PDF: %v\n", err) // デバッグ用ログ
		return
	}

	fmt.Fprintf(w, "ドキュメント '%s' をプリンター '%s' に正常に送信しました。", handler.Filename, printerName)
	log.Println("PDF印刷リクエストが正常に処理されました。")                    // ログ出力
	fmt.Println("PDF print request processed successfully.") // デバッグ用ログ
	// 注意: 印刷プロセスはバックグラウンドで実行されるため、このハンドラ内では一時ファイルを安全に削除できません。
	// ファイルはOSの一時ディレクトリに残りますが、これは意図した動作です。
}

// printPDF は指定されたPDFファイルを指定されたプリンターに印刷します。
// Adobe Acrobat Reader DC (AcroRd32.exe) を使用することを想定しています。
func printPDF(documentPath, printerName string) error {
	// 注意: この関数は印刷コマンドをバックグラウンドで開始し、完了を待ちません。
	// これにより、HTTPハンドラがブロックされるのを防ぎます。
	// Adobe Acrobat Reader DC の実行可能ファイルのパス。
	// 環境によって異なる場合があります。必要に応じて変更してください。
	// 例: "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe"
	// または、PATH環境変数にAdobe Readerのパスが追加されている場合は "AcroRd32.exe" のみでも動作する場合があります。
	adobeReaderPath := os.Getenv("ADOBE_READER_PATH")
	if adobeReaderPath == "" {
		// adobeReaderPath = "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe" // デフォルトパス
		adobeReaderPath = "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe" // デフォルトパス
		log.Printf("ADOBE_READER_PATH 環境変数が設定されていません。デフォルトパスを使用します: %s\n", adobeReaderPath)
		fmt.Printf("ADOBE_READER_PATH environment variable not set. Using default path: %s\n", adobeReaderPath) //
	} else {
		log.Printf("環境変数からAdobe Readerのパスを使用しています: %s\n", adobeReaderPath)
		fmt.Printf("Using Adobe Reader path from environment variable: %s\n", adobeReaderPath) // デバッグ用ログ
	}

	// Adobe Readerが存在するか確認
	if _, err := os.Stat(adobeReaderPath); os.IsNotExist(err) {
		// PATHからAcrobat.exeを探す（もしPATHに追加されていれば）
		cmdPath, err := exec.LookPath("Acrobat.exe")
		if err != nil {
			log.Printf("Adobe Acrobat Reader (Acrobat.exe) が '%s' または PATH に見つかりませんでした: %v", adobeReaderPath, err)
			return fmt.Errorf("Adobe Acrobat Reader (Acrobat.exe) が '%s' または PATH に見つかりませんでした: %w", adobeReaderPath, err)
		}
		adobeReaderPath = cmdPath // PATHで見つかったパスを使用
	}

	// 印刷コマンドを構築します。
	// /t: 印刷ダイアログを表示せずに指定されたプリンターにファイルを印刷します。
	// /h: アプリケーションウィンドウを非表示にします。
	// cmd := exec.Command(adobeReaderPath, "/t", documentPath, printerName)
	quotedAdobeReaderPath := fmt.Sprintf(`"%s"`, adobeReaderPath)

	// プリンター名にスペースや特殊文字が含まれる場合に備え、明示的に引用符で囲む。
	quotedPrinterName := fmt.Sprintf(`"%s"`, printerName)

	// documentPath も明示的に引用符で囲む。
	quotedDocumentPath := fmt.Sprintf(`"%s"`, documentPath)
	cmd := exec.Command(quotedAdobeReaderPath, "/t", quotedDocumentPath, quotedPrinterName) // すべて引用符付きの引数を渡す

	log.Printf("印刷コマンドを構築しました: %s %s %s %s", cmd.Args[0], cmd.Args[1], cmd.Args[2], cmd.Args[3])           // ログ出力
	log.Printf("印刷コマンドを実行しています: %s %s %s %s", adobeReaderPath, "/t", documentPath, printerName)            // ログ出力
	fmt.Printf("Executing print command: %s %s %s %s\n", adobeReaderPath, "/t", documentPath, printerName) // デバッグ用ログ

	// cmd.Run() はコマンドが完了するまでブロックし、GUIアプリケーションがハングするとレスポンスが返らなくなる原因になります。
	// cmd.Start() はコマンドを非同期に開始し、すぐに処理を返すため、HTTPハンドラは即座にレスポンスを返すことができます。
	err := cmd.Start()
	if err != nil {
		log.Printf("コマンドの開始に失敗しました: %v", err)
		return fmt.Errorf("コマンドの開始に失敗しました: %w", err)
	}

	// プロセスが正常に開始されたことをログに記録します。コマンドの完了は待ちません。
	log.Printf("印刷コマンドがバックグラウンドで正常に開始されました (PID: %d)", cmd.Process.Pid)
	fmt.Printf("Print command started successfully in background (PID: %d).\n", cmd.Process.Pid)

	// Goプログラムが子プロセスの終了を待たないように、プロセスを「解放」します。
	// これにより、Acrobat.exeは完全に独立して実行され、Go側のHTTPハンドラがブロックされるリスクをさらに低減します。
	err = cmd.Process.Release()
	if err != nil {
		// Releaseの失敗は致命的ではない可能性が高いため、警告ログに記録するに留めます。
		log.Printf("警告: プロセスの解放に失敗しました (PID: %d): %v", cmd.Process.Pid, err)
	}
	return nil
}

// main 関数はプログラムのエントリポイントです。
func main() {
	fmt.Println("Entering main function.") // デバッグ用: main関数開始をコンソールに出力
	log.Println("アプリケーションを開始します。")         // ログ出力

	// 多重起動をチェックし、古いプロセスを終了させる
	handleMultipleInstances()

	// systrayを開始し、onReadyとonExit関数を登録します。
	systray.Run(onReady, onExit)
}

// handleMultipleInstances は、同じアプリケーションの他のインスタンスを探して終了させます。
func handleMultipleInstances() {
	// この機能はWindowsでのみ動作します。
	if runtime.GOOS != "windows" {
		return
	}

	// 現在の実行ファイル名を取得します。
	exe, err := os.Executable()
	if err != nil {
		// 致命的ではないため、警告ログを出力して続行します。
		log.Printf("警告: 実行ファイルパスの取得に失敗しました: %v", err)
		return
	}
	exeName := filepath.Base(exe)
	currentPid := os.Getpid()

	// tasklistコマンドで同じ名前のプロセスを検索します。
	cmd := exec.Command("tasklist", "/FI", fmt.Sprintf("IMAGENAME eq %s", exeName), "/FO", "CSV", "/NH")
	output, err := cmd.Output()
	if err != nil {
		log.Printf("警告: 既存プロセスの検索に失敗しました: %v", err)
		return
	}

	// CSV出力をパースして自分以外のプロセスをkillします。
	scanner := bufio.NewScanner(bytes.NewReader(output))
	for scanner.Scan() {
		line := scanner.Text()
		// CSVの各フィールドはダブルクォートで囲まれています。
		// 例: "http-printer.exe","12345","Console","1","12,345 K"
		fields := strings.Split(line, `","`)
		if len(fields) < 2 {
			continue
		}

		// PIDは2番目のフィールドです。
		pid, err := strconv.Atoi(fields[1])
		if err != nil {
			continue
		}

		// 自分自身のプロセスでなければ終了させます。
		if pid != currentPid {
			log.Printf("既存のプロセス (PID: %d) が見つかりました。終了を試みます。", pid)
			if proc, err := os.FindProcess(pid); err == nil {
				if err := proc.Kill(); err != nil {
					log.Printf("警告: プロセス (PID: %d) の終了に失敗しました: %v", pid, err)
				} else {
					log.Printf("プロセス (PID: %d) を正常に終了しました。", pid)
				}
			}
		}
	}
}

// onReady はタスクトレイアイコンが準備できたときに呼び出されます。
func onReady() {
	systray.SetIcon(IconData) // icon.goで定義されたアイコンデータを設定

	systray.SetTitle("Go HTTP Printer")             // タイトルを設定
	systray.SetTooltip("Go HTTP Printer (ポート8080)") // ツールチップを設定

	// HTTPサーバーを新しいゴルーチンで起動します。
	go startHTTPServer()

	// メニュー項目を追加
	mOpen := systray.AddMenuItem("ブラウザで開く", "ブラウザでローカルサーバーを開きます")
	mQuit := systray.AddMenuItem("終了", "アプリケーションを終了します")

	// メニュー項目のクリックイベントを処理
	for {
		select {
		case <-mOpen.ClickedCh:
			log.Println("ブラウザで開くがクリックされました。")
			openbrowser("http://localhost:8080")
		case <-mQuit.ClickedCh:
			log.Println("終了がクリックされました。アプリケーションを終了します。")
			systray.Quit() // systrayを終了し、onExitをトリガーします
			return
		}
	}
}

// onExit はアプリケーションが終了するときに呼び出されます。
func onExit() {
	log.Println("アプリケーションが終了しました。")
	// 必要に応じてクリーンアップ処理を追加
	os.Exit(0) // プログラムを正常終了
}

// openbrowser は指定されたURLをデフォルトのウェブブラウザで開きます。
func openbrowser(url string) {
	var err error
	switch runtime.GOOS {
	case "linux":
		err = exec.Command("xdg-open", url).Start()
	case "windows":
		err = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
	case "darwin":
		err = exec.Command("open", url).Start()
	default:
		err = fmt.Errorf("サポートされていないOSです")
	}
	if err != nil {
		log.Printf("ブラウザを開けませんでした: %v", err)
	}
}
