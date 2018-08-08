package main

import (
	"bytes"
	"fmt"
	"github.com/gin-gonic/gin"
	"github.com/tealeg/xlsx"
	"net/http"
	"strconv"
)

func main() {
	r := gin.Default()
	r.GET("/ping", func(c *gin.Context) {

		genexcel()
		c.JSON(200, gin.H{
			"message": "pong",
		})
	})

	r.GET("/download", ginexcel)
	r.Run() // listen and serve on 0.0.0.0:8080
}

type ContactsModel struct {
	ID     int
	Name   string
	Number string
	Age    int
}

func getContacts() []ContactsModel {

	rs := make([]ContactsModel, 100)
	for i := 0; i < 100; i++ {
		c := ContactsModel{}
		c.ID = i
		c.Name = fmt.Sprintf("test-%d", i)
		c.Number = fmt.Sprintf("format %d data", i)
		c.Age = i % 10

		rs[i] = c
	}

	return rs
}

// 导出到浏览器
func ginexcel(ctx *gin.Context) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var err error
	var contacts []ContactsModel

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	contacts = getContacts()

	row = sheet.AddRow()
	row.AddCell().Value = "Name"
	row.AddCell().Value = "Number"
	row.AddCell().Value = "Age"
	for _, contact := range contacts {
		row = sheet.AddRow()
		row.AddCell().Value = contact.Name
		row.AddCell().Value = contact.Number
		row.AddCell().Value = strconv.Itoa(contact.Age)
	}
	buf := new(bytes.Buffer)
	err = file.Write(buf)
	if err != nil {
		fmt.Printf(err.Error())
		return
	}
	ctx.Header("Content-Description", "File Transfer")
	ctx.Header("Content-Disposition", "attachment; filename=contacts.xlsx")
	ctx.Data(http.StatusOK, "text/xlsx", buf.Bytes())
}

// 导出到文件
func genexcel() {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "I am a cell!"
	err = file.Save("MyXLSXFile.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}
