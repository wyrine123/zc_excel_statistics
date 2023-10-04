package main

import (
	"bytes"
	"context"
	"embed"
	"fmt"
	"github.com/gin-gonic/gin"
	"github.com/prometheus/client_golang/prometheus/promhttp"
	"gopkg.in/yaml.v3"
	"html/template"
	"io"
	"log"
	"log/slog"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"os/signal"
	"runtime"
	"runtime/debug"
	"strings"
	"syscall"
	"time"
	"zc_excel_statistics/internal"
)

type ExcelConfig struct {
	Sign   ExcelCommonConfig `yaml:"sign"`
	Design ExcelCommonConfig `yaml:"design"`
}

type ExcelCommonConfig struct {
	IgnoreExcelLastRow      bool     `yaml:"ignore_excel_last_row"`
	ContractMoneyColumnName string   `yaml:"contract_money_column_name"`
	DesignerColumnNameList  []string `yaml:"designer_column_name_list"`
}

//go:embed configs/excel.yaml
//go:embed web
var f embed.FS

func main() {
	// 加载配置
	// 初始化log配置
	handler := slog.NewJSONHandler(os.Stdout, nil)
	logger := slog.New(handler)

	slog.SetDefault(logger)

	buildInfo, _ := debug.ReadBuildInfo()

	slog.Info("test", "build_info", buildInfo)

	// 解析配置文件
	configs, err := f.ReadFile("configs/excel.yaml")
	if err != nil {
		panic("读取配置文件失败: " + err.Error())
	}

	var excelConfig ExcelConfig
	decoder := yaml.NewDecoder(bytes.NewReader(configs))
	err = decoder.Decode(&excelConfig)
	if err != nil {
		panic("解析配置文件失败: " + err.Error())
	}

	slog.Info("excel_config", "config", excelConfig)

	ctx, stop := signal.NotifyContext(context.Background(), syscall.SIGINT, syscall.SIGTERM)
	defer stop()

	r := gin.Default()

	temp := template.Must(template.New("*").ParseFS(f, "web/*.html"))
	r.SetHTMLTemplate(temp)

	r.GET("/metrics", gin.WrapH(promhttp.Handler()))

	// 前端

	r.GET("", func(c *gin.Context) {
		c.HTML(http.StatusOK, "index.html", nil)
	})

	// 服务端处理
	r.POST("/sign/receipt", func(c *gin.Context) {
		var err1 error
		excelFile, oriFileName, err1 := internal.ReadExcelFileFromHttpMultipartFileHeader(c)
		if err1 != nil {
			c.JSON(http.StatusBadRequest, "文件上传失败")
			return
		}

		sign := excelConfig.Sign
		newExcelBytes, err1 := internal.CalExcel(
			excelFile,
			sign.IgnoreExcelLastRow,
			sign.ContractMoneyColumnName,
			sign.DesignerColumnNameList,
		)
		if err1 != nil {
			c.JSON(http.StatusBadRequest, err1.Error())
			return
		}

		slog.Info(fmt.Sprintf("原始文件名称：%s", oriFileName))
		fileName := oriFileName
		fileName = strings.Replace(fileName, ".xlsx", "-统计.xlsx", 1)
		slog.Info(fmt.Sprintf("新文件名称：%s", fileName))

		c.Writer.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8")
		c.Writer.Header().Set("Content-Disposition", fmt.Sprintf("filename=%s", url.QueryEscape(fileName)))
		rr := bytes.NewReader(newExcelBytes.Bytes())
		_, _ = io.Copy(c.Writer, rr)
	}) // 签单

	r.POST("/design/receipt", func(c *gin.Context) {
		var err1 error
		excelFile, oriFileName, err1 := internal.ReadExcelFileFromHttpMultipartFileHeader(c)
		if err1 != nil {
			c.JSON(http.StatusBadRequest, "文件上传失败")
			return
		}

		design := excelConfig.Design
		newExcelBytes, err1 := internal.CalExcel(
			excelFile,
			design.IgnoreExcelLastRow,
			design.ContractMoneyColumnName,
			design.DesignerColumnNameList,
		)
		if err1 != nil {
			c.JSON(http.StatusBadRequest, err1.Error())
			return
		}

		slog.Info(fmt.Sprintf("原始文件名称：%s", oriFileName))
		fileName := oriFileName
		fileName = strings.Replace(fileName, ".xlsx", "-统计.xlsx", 1)
		slog.Info(fmt.Sprintf("新文件名称：%s", fileName))

		c.Writer.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8")
		c.Writer.Header().Set("Content-Disposition", fmt.Sprintf("filename=%s", url.QueryEscape(fileName)))
		rr := bytes.NewReader(newExcelBytes.Bytes())
		_, _ = io.Copy(c.Writer, rr)
	}) // 设计费

	srv := &http.Server{
		Addr:     ":8080",
		Handler:  r,
		ErrorLog: slog.NewLogLogger(handler, slog.LevelInfo),
	}

	go func() {
		if err := srv.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			log.Fatalf("listen: %s\n", err)
		}
	}()

	go func() {
		<-time.After(time.Second * 1)

		var cmd *exec.Cmd
		switch runtime.GOOS {
		case "windows":
			// Windows 上使用 start 命令来打开默认浏览器
			cmd = exec.Command("cmd", "/c", "start", "http://127.0.0.1:8080/")
		case "darwin":
			// macOS 上使用 open 命令来打开默认浏览器
			cmd = exec.Command("open", "http://127.0.0.1:8080/")
		default:
			// Linux 和其他系统上使用 xdg-open 命令来打开默认浏览器
			cmd = exec.Command("xdg-open", "http://127.0.0.1:8080/")
		}

		err1 := cmd.Run()
		if err1 != nil {
			slog.Error("无法打开浏览器:", "err", err1)
		}
	}()

	<-ctx.Done()
	stop()
	log.Println("shutting down gracefully, press Ctrl+C again to force")

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()

	if err := srv.Shutdown(ctx); err != nil {
		log.Fatal("Server forced to shutdown: ", err)
	}
	log.Println("Server exiting")
}
