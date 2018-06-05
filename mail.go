package main

import (
	"database/sql"
	"errors"
	"fmt"
	"time"

	"github.com/Unknwon/goconfig"
	"github.com/jasonlvhit/gocron"
	_ "github.com/lib/pq"
	"github.com/tealeg/xlsx"
	"gopkg.in/gomail.v2"
)

var cfg *goconfig.ConfigFile

// var db *sql.DB

func main() {

	gocron.Every(1).Day().At("09:00").Do(func() {
		//fmt.Println(1)
		getSQLGroup()
	})

	<-gocron.Start()

	//sendmail()

	// getSQLGroup()
}

func init() {
	var err error
	cfg, err = goconfig.LoadConfigFile("./conf.ini")
	if err != nil {
		fmt.Println(err)
	}
	// err = openDb()
	// if err != nil {
	// 	fmt.Println(err)
	// }
}

func openDb() (db *sql.DB, err error) {

	// if db != nil || db.Stats() !=  {
	// 	return
	// }

	conn := cfg.MustValue("db", "conn")
	if conn == "" {
		err = errors.New("数据库链接异常")
		return
	}
	db, err = sql.Open("postgres", conn)
	return
}

//获取Sql节点的所有Key，并循环
func getSQLGroup() {
	keys := cfg.GetKeyList("sql")
	for _, v := range keys {
		sqlStr := cfg.MustValue("sql", v)
		mailAry := cfg.MustValueArray("mailto", v, ",")
		fileName := attachment(sqlStr, v)
		sendmail(mailAry, fileName)
	}
}

func query(strSQL string, fn func([][]byte)) {

	db, err := openDb()
	if err != nil {
		fmt.Println(err)
		return
	}
	defer db.Close()

	//strSQL := cfg.MustValue("sql", "group1")

	//统计昨天的数据
	yestday := time.Now().AddDate(0, 0, -1).Format("2006-01-02")

	rows, _ := db.Query(strSQL, yestday)
	defer rows.Close()

	cls, _ := rows.Columns()

	vals := make([][]byte, len(cls))
	scans := make([]interface{}, len(cls))
	for i := range cls {
		scans[i] = &vals[i]
	}
	for rows.Next() {
		err := rows.Scan(scans...)
		if err != nil {
			fmt.Println(err)
			continue
		}
		fn(vals)
	}
}

func attachment(strSQL, groupName string) string {
	xlsFile := xlsx.NewFile()
	sht, err := xlsFile.AddSheet("sheet1")
	if err != nil {
		fmt.Println(err)
		return ""
	}
	query(strSQL, func(buf [][]byte) {
		row := sht.AddRow()
		for _, c := range buf {
			cell := row.AddCell()
			cell.SetValue(string(c))
		}
	})
	//buf := bytes.NewBuffer([]byte{})
	filename := fmt.Sprintf("./file/%s-%s.xlsx", time.Now().Format("20060102150405"), groupName)
	xlsFile.Save(filename)
	fmt.Println("文件生成完成...")
	return filename
}

func sendmail(to []string, attachment string) {

	from := cfg.MustValue("smtp", "from")

	msg := gomail.NewMessage()
	msg.SetHeader("From", from)

	//to := cfg.MustValueArray("mailto", "group1", ",")

	msg.SetHeader("To", to...)

	//attachment := attachment()

	subject := fmt.Sprintf("%s-%s", cfg.MustValue("mail", "subject"), time.Now().AddDate(0, 0, -1).Format("2006-01-02"))

	msg.SetHeader("Subject", subject)

	if attachment == "" {
		msg.SetBody("text/html", "<color='red'>报表生成异常</color>")
	} else {
		msg.SetBody("text/html", "见附件")
		msg.Attach(attachment)
	}

	d := gomail.NewDialer(cfg.MustValue("smtp", "host"), cfg.MustInt("smtp", "port"), from, cfg.MustValue("smtp", "password"))

	if err := d.DialAndSend(msg); err != nil {
		fmt.Println(err)
	}

}
