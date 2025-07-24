package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"time"

	"bytes"         // bytes.Bufferのために追加
	"os/exec"       // 外部コマンド実行のために追加
	"path/filepath" // パス操作のために追加

	"golang.org/x/sys/windows/svc"
	"golang.org/x/sys/windows/svc/debug"
	"golang.org/x/sys/windows/svc/eventlog"
)

// elog はイベントログへの書き込みに使用されます。
var elog debug.Log

// myService はWindowsサービスとしての振る舞いを定義する構造体です。
type myService struct{}

// Execute はサービスが開始されたときに呼び出されるメイン関数です。
// サービスの状態変更リクエストを処理し、HTTPサーバーを起動します。
func (m *myService) Execute(args []string, r <-chan svc.ChangeRequest, status chan<- svc.Status) (ssec bool, errno uint32) {
	// サービスが受け入れるコマンド（停止、シャットダウン）を定義します。
	const cmdsAccepted = svc.AcceptStop | svc.AcceptShutdown

	// サービスの状態を「開始保留中」に設定します。
	status <- svc.Status{State: svc.StartPending}
	// イベントログにサービス開始の情報を記録します。
	elog.Info(1, "Service starting (Execute function).")

	// HTTPサーバーを新しいゴルーチンで起動します。
	go startHTTPServer()

	// サービスの状態を「実行中」に設定し、受け入れるコマンドを通知します。
	status <- svc.Status{State: svc.Running, Accepts: cmdsAccepted}

loop:
	// サービスコントロールマネージャーからのリクエストを待ちます。
	for c := range r {
		switch c.Cmd {
		case svc.Interrogate:
			// サービスの状態照会リクエストに応答します。
			elog.Info(1, "Service Interrogate request received.")
			status <- c.CurrentStatus // 現在の状態を返します
		case svc.Stop, svc.Shutdown:
			// 停止またはシャットダウンリクエストを受け取った場合、ループを抜けてサービスを終了します。
			elog.Info(1, "Service stop/shutdown request received. Exiting loop.")
			break loop
		default:
			// 未知の制御リクエストを受け取った場合、エラーを記録します。
			elog.Error(1, fmt.Sprintf("Unexpected control request #%d", c.Cmd))
		}
	}
	// サービスの状態を「停止保留中」に設定します。
	status <- svc.Status{State: svc.StopPending}
	// イベントログにサービス終了の情報を記録します。
	elog.Info(1, "Service finished.")
	return
}

// startHTTPServer はHTTPサーバーを起動する関数です。
// サービスとして実行されていない場合は、コンソールに直接ログを出力します。
func startHTTPServer() {
	fmt.Println("Entering startHTTPServer function.") // デバッグ用: 関数開始をコンソールに出力

	// ルートパス ("/") へのリクエストを処理するハンドラを設定します。
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		fmt.Fprintf(w, "Hello from Go HTTP Server running as a Windows Service!")
	})

	// PDF印刷用の新しいハンドラを追加
	http.HandleFunc("/print-pdf", printPDFHandler)
	fmt.Println("Added /print-pdf handler.") // デバッグ用ログ

	// HTTPサーバーがリッスンするポートを設定します。
	port := ":8080"
	fmt.Printf("Attempting to start HTTP server on port %s\n", port) // デバッグ用: ポート情報をコンソールに出力

	// HTTPサーバーを起動し、エラーがあれば処理します。
	if err := http.ListenAndServe(port, nil); err != nil {
		// HTTPサーバーの起動に失敗した場合、エラーメッセージをコンソールに出力します。
		fmt.Printf("HTTP server failed: %v\n", err)
		// プログラムを終了します。これにより、エラーがすぐに確認できます。
		os.Exit(1)
	}
	fmt.Println("HTTP server started successfully.") // デバッグ用: サーバー起動成功をコンソールに出力
}

// printPDFHandler は /print-pdf エンドポイントのリクエストを処理します。
// POST multipart/form-data のボディからPDFファイルとプリンター名を受け取り、PDFを印刷します。
func printPDFHandler(w http.ResponseWriter, r *http.Request) {
	fmt.Println("Received request for /print-pdf.")  // デバッグ用ログ
	elog.Info(1, "Received request for /print-pdf.") // イベントログにも出力

	// POSTメソッド以外は許可しない
	if r.Method != http.MethodPost {
		http.Error(w, "Only POST method is allowed for this endpoint.", http.StatusMethodNotAllowed)
		fmt.Println("Error: Only POST method is allowed.") // デバッグ用ログ
		return
	}

	// multipart/form-data をパースします。最大10MBのファイルを受け入れます。
	err := r.ParseMultipartForm(10 << 20) // 10MB
	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to parse multipart form: %v", err), http.StatusBadRequest)
		fmt.Printf("Error: Failed to parse multipart form: %v\n", err) // デバッグ用ログ
		return
	}

	// プリンター名を取得します。
	printerName := r.FormValue("printer")
	if printerName == "" {
		http.Error(w, "Missing 'printer' parameter in form data.", http.StatusBadRequest)
		fmt.Println("Error: Missing 'printer' parameter.") // デバッグ用ログ
		return
	}

	// アップロードされたPDFファイルを取得します。
	file, handler, err := r.FormFile("document")
	if err != nil {
		http.Error(w, fmt.Sprintf("Missing 'document' file in form data: %v", err), http.StatusBadRequest)
		fmt.Printf("Error: Missing 'document' file: %v\n", err) // デバッグ用ログ
		return
	}
	defer file.Close() // 関数終了時にファイルを閉じます。

	// 一時ファイルを作成し、アップロードされたPDFを書き込みます。
	// サービス実行アカウントが書き込み権限を持つ一時ディレクトリを使用します。
	// tempDir := os.TempDir() // システムの一時ディレクトリ
	//tempDir にcurrent directoryを使用する場合
	tempDir, err := os.Getwd() // 現在の作業ディレクトリ
	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to get current directory: %v", err), http.StatusInternalServerError)
		fmt.Printf("Error: Failed to get current directory: %v\n", err) //
		return
	}

	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to get current directory: %v", err), http.StatusInternalServerError)
		fmt.Printf("Error: Failed to get current directory: %v\n", err) // デバッグ用ログ
		return
	}
	tempFilePath := filepath.Join(tempDir, handler.Filename)
	fmt.Printf("Saving uploaded file to temporary path: %s\n", tempFilePath)                          // デバッグ用ログ
	elog.Info(1, fmt.Sprintf("Attempting to save uploaded file to temporary path: %s", tempFilePath)) // イベントログに追加

	tempFile, err := os.Create(tempFilePath)
	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to create temporary file: %v", err), http.StatusInternalServerError)
		fmt.Printf("Error: Failed to create temporary file: %v\n", err) // デバッグ用ログ
		return
	}
	defer tempFile.Close() // 関数終了時にファイルを閉じます。
	// defer os.Remove(tempFilePath) // 印刷後に一時ファイルを削除します。

	_, err = io.Copy(tempFile, file)
	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to save uploaded file: %v", err), http.StatusInternalServerError)
		fmt.Printf("Error: Failed to save uploaded file: %v\n", err)                                     // デバッグ用ログ
		elog.Error(1, fmt.Sprintf("Error: Failed to save uploaded file to '%s': %v", tempFilePath, err)) // イベントログに追加

		return
	}
	tempFile.Close()                                                                    // ここで明示的にファイルを閉じる
	elog.Info(1, fmt.Sprintf("Successfully saved to temporary file: %s", tempFilePath)) // イベントログに追加

	fmt.Printf("Attempting to print document '%s' to printer '%s'.\n", tempFilePath, printerName) // デバッグ用ログ

	// PDF印刷を実行します。
	err = printPDF(tempFilePath, printerName)
	if err != nil {
		http.Error(w, fmt.Sprintf("Failed to print PDF: %v", err), http.StatusInternalServerError)
		fmt.Printf("Error printing PDF: %v\n", err) // デバッグ用ログ
		// サービスとして実行されている場合、イベントログにもエラーを記録
		if elog != nil {
			elog.Error(1, fmt.Sprintf("Failed to print PDF '%s' to '%s': %v", tempFilePath, printerName, err))
		}
		return
	}

	fmt.Fprintf(w, "Successfully sent document '%s' to printer '%s'.", handler.Filename, printerName)
	fmt.Println("PDF print request processed successfully.") // デバッグ用ログ
	// ★印刷成功時にのみ一時ファイルを削除する★
	err = os.Remove(tempFilePath)
	if err != nil {
		elog.Error(1, fmt.Sprintf("Failed to remove temporary file '%s': %v", tempFilePath, err))
	}
}

// printPDF は指定されたPDFファイルを指定されたプリンターに印刷します。
// Adobe Acrobat Reader DC (AcroRd32.exe) を使用することを想定しています。
func printPDF(documentPath, printerName string) error {
	// Adobe Acrobat Reader DC の実行可能ファイルのパス。
	// 環境によって異なる場合があります。必要に応じて変更してください。
	// 例: "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe"
	// または、PATH環境変数にAdobe Readerのパスが追加されている場合は "AcroRd32.exe" のみでも動作する場合があります。
	adobeReaderPath := os.Getenv("ADOBE_READER_PATH")
	if adobeReaderPath == "" {
		adobeReaderPath := "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe"                                      // デフォルトパス
		fmt.Printf("ADOBE_READER_PATH environment variable not set. Using default path: %s\n", adobeReaderPath)              //
		elog.Info(1, fmt.Sprintf("ADOBE_READER_PATH environment variable not set. Using default path: %s", adobeReaderPath)) // イベントログに追加
	} else {
		fmt.Printf("Using Adobe Reader path from environment variable: %s\n", adobeReaderPath)              // デバッグ用ログ
		elog.Info(1, fmt.Sprintf("Using Adobe Reader path from environment variable: %s", adobeReaderPath)) // イベントログに追加
	}

	// Adobe Readerが存在するか確認
	if _, err := os.Stat(adobeReaderPath); os.IsNotExist(err) {
		// PATHからAcroRd32.exeを探す（もしPATHに追加されていれば）
		cmdPath, err := exec.LookPath("Acrobat.exe")
		if err != nil {
			elog.Error(1, fmt.Sprintf("Adobe Acrobat Reader (Acrobat.exe) not found at '%s' or in PATH: %v", adobeReaderPath, err)) // イベントログに追加
			return fmt.Errorf("adobe Acrobat Reader (Acrobat.exe) not found at '%s' or in PATH: %w", adobeReaderPath, err)
		}
		adobeReaderPath = cmdPath // PATHで見つかったパスを使用
	}

	// 印刷コマンドを構築します。
	// /t: 印刷ダイアログを表示せずに指定されたプリンターにファイルを印刷します。
	// /h: アプリケーションウィンドウを非表示にします。
	cmd := exec.Command(adobeReaderPath, "/t", documentPath, printerName)
	// ここが重要！外部コマンドの標準出力と標準エラー出力をキャプチャする
	var stdoutBuf, stderrBuf bytes.Buffer
	cmd.Stdout = &stdoutBuf
	cmd.Stderr = &stderrBuf

	// コマンドの標準出力と標準エラー出力をキャプチャしてデバッグに役立てることもできますが、
	// ここではシンプルに実行します。
	// cmd.Stdout = os.Stdout
	// cmd.Stderr = os.Stderr
	elog.Info(1, fmt.Sprintf("Executing print command: %s %s %s %s", adobeReaderPath, "/t", documentPath, printerName)) // イベントログに追加
	fmt.Printf("Executing print command: %s %s %s %s\n", adobeReaderPath, "/t", documentPath, printerName)              // デバッグ用ログ

	// コマンドを実行し、完了を待ちます。
	err := cmd.Run()

	// コマンド実行後のStdout/Stderrをイベントログに出力
	if stdoutBuf.Len() > 0 {
		elog.Info(1, fmt.Sprintf("Print command Stdout: %s", stdoutBuf.String()))
		fmt.Printf("Print command Stdout: %s\n", stdoutBuf.String()) // コンソールにも
	}
	if stderrBuf.Len() > 0 {
		elog.Error(1, fmt.Sprintf("Print command Stderr: %s", stderrBuf.String())) // ここにエラーが出れば、その内容を！
		fmt.Printf("Print command Stderr: %s\n", stderrBuf.String())               // コンソールにも
	}

	if err != nil {
		elog.Error(1, fmt.Sprintf("Command execution failed: %v", err))
		return fmt.Errorf("command execution failed: %w", err)
	}

	fmt.Println("Print command executed successfully.") // デバッグ用ログ
	return nil
}

// main 関数はプログラムのエントリポイントです。
// サービスとして実行されているか、インタラクティブセッションで実行されているかを判断します。
func main() {
	fmt.Println("Entering main function.") // デバッグ用: main関数開始をコンソールに出力

	// 現在のセッションがWindowsサービスであるかどうかを判断します。
	isWindowsService, err := svc.IsWindowsService() // 変数名をより明確に
	if err != nil {
		// 判断に失敗した場合、致命的なエラーとしてログに出力し、プログラムを終了します。
		// elogはまだ開かれていない可能性があるので、log.Fatalfを使用します。
		log.Fatalf("Failed to determine if we are running as a Windows service: %v", err)
	}
	// デバッグ用: セッションタイプをコンソールに出力
	fmt.Printf("Is Windows Service: %t\n", isWindowsService) // サービスかどうかを直接表示

	if isWindowsService { // 論理を修正: サービスとして実行されている場合
		// サービスとして実行する場合の処理
		fmt.Println("Running as a Windows service.") // デバッグ用: サービスとして実行中であることをコンソールに出力

		// イベントログを開き、サービスからのログメッセージを記録できるようにします。
		elog, err = eventlog.Open("MyGoHTTPServer") // イベントログの名前
		if err != nil {
			// イベントログのオープンに失敗した場合、致命的なエラーとしてログに出力し、プログラムを終了します。
			log.Fatalf("Failed to open event log: %v", err)
		}
		defer elog.Close() // 関数終了時にイベントログを閉じます。

		// イベントログにサービス開始の情報を記録します。
		elog.Info(1, "Service is about to run...")
		// サービスを実行します。この呼び出しはサービスが停止するまでブロックされます。
		err = svc.Run("MyGoHTTPServer", &myService{}) // サービスの名前とサービスハンドラを渡します。
		if err != nil {
			// サービス実行中にエラーが発生した場合、イベントログにエラーを記録します。
			elog.Error(1, fmt.Sprintf("Service failed: %v", err))
		}
		// サービスが停止したことをイベントログに記録します。
		elog.Info(1, "Service stopped.")
		return // サービスとしての実行が完了したら、プログラムを終了します。
	} else {
		// インタラクティブセッション（通常のコンソールアプリケーション）で実行する場合の処理
		fmt.Println("Running in interactive session (for debugging). Press Ctrl+C to stop.")
		// HTTPサーバーを起動します。
		startHTTPServer()
		// プロセスが終了しないように無限ループで待機します。
		// (startHTTPServerがエラーでos.Exit(1)しない限りここに到達します)
		for {
			time.Sleep(time.Second)
		}
	}
}

// installService はWindowsサービスをインストールするための補助関数です。
// 通常、この関数は別のプログラムまたはスクリプトから呼び出されます。
func installService(name, desc string) error {
	exepath := os.Args[0] // 現在の実行可能ファイルのパスを取得します。
	// sc create コマンドを構築します。
	cmd := fmt.Sprintf("sc create %s binPath= \"%s\" start= auto DisplayName= \"%s\"", name, exepath, desc)
	fmt.Printf("Running command: %s\n", cmd)
	// TODO: os/execパッケージを使用して直接実行することも可能ですが、ここではコマンド文字列を表示するだけです。
	return nil // エラーハンドリングを追加してください。
}

// uninstallService はWindowsサービスをアンインストールするための補助関数です。
// 通常、この関数は別のプログラムまたはスクリプトから呼び出されます。
func uninstallService(name string) error {
	// sc delete コマンドを構築します。
	cmd := fmt.Sprintf("sc delete %s", name)
	fmt.Printf("Running command: %s\n", cmd)
	return nil // エラーハンドリングを追加してください。
}
