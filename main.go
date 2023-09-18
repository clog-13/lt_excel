package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	Init()
}

func Init() {
	installation := excelize.NewFile()
	defer func() {
		if err := installation.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	index, err := installation.NewSheet("data sheet")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := installation.SaveAs("source/端口设备.xlsx"); err != nil {
			fmt.Println(err)
		}
	}()
	installation.SetActiveSheet(index)

	installation_title := [...]string{"所属机房名称**", "安装位置", "设备名称**", "设备编码**", "设备型号", "生产厂商", "分光比**", "分光级别**", "所属OLT设备**", "所属OLT设备端口**", "上联OBD设备", "上联OBD设备端口", "维护方式", "维护单位名称", "工程项目", "竣工日期", "施工单位", "所属区域**", "所属局向", "是否是自激活设备**", "二维码", "经度", "纬度", "建设模式*", "合作模式", "产权归属时限", "是否虚拟资源", "错误信息", "错误编码"}
	title_no := [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"}
	for idx, val := range installation_title {
		installation.SetCellValue("data sheet", title_no[idx]+"1", val)
	}

	// TODO: ReadOnly
	source, err := excelize.OpenFile("source/1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// calculate installation count
	cols, _ := source.GetCols("号线资源表")
	SECOND_BEAM_SPLITTER := 1
	for i := 3; ; i++ {
		val := cols[0][i]
		if err != nil {
			fmt.Println(err)
			return
		}
		if strings.TrimSpace(val) == "" {
			break
		} else {
			num, err := strconv.Atoi(val)
			if err != nil {
				fmt.Println("Atoi err:", err)
				return
			}
			if num > SECOND_BEAM_SPLITTER {
				SECOND_BEAM_SPLITTER = num
			}
		}
	}
	FRIST_BEAM_SPLITTER := SECOND_BEAM_SPLITTER / 8
	if SECOND_BEAM_SPLITTER%8 != 0 {
		FRIST_BEAM_SPLITTER++
	}

	// OBD设备
	// TODO: 所属机房，设备编码，所属OLT设备
	// ==================================================================================================

	// 所属机房名称** || 设备名称
	pre_fbs := ""
	fbs_cnt, sbs_cnt := 1, 1
	for i := 0; i < SECOND_BEAM_SPLITTER; i++ {
		// 设备名称，一级分光 || 安装位置，一级分光
		cur, _ := source.GetCellValue("号线资源表", "I"+strconv.Itoa(i+3))
		if pre_fbs != cur {
			pre_fbs = cur
			// TODO: 文字切割
			installation.SetCellValue("data sheet", "B"+strconv.Itoa(fbs_cnt+1), cur[:len(cur)-70]+strconv.Itoa(fbs_cnt)+"号分光器")
			olt, _ := source.GetCellValue("号线资源表", "P"+strconv.Itoa(i+3))
			parts := strings.Split(olt, "-")
			installation.SetCellValue("data sheet", "C"+strconv.Itoa(fbs_cnt+1), cur[:len(cur)-70]+"OBD0"+parts[len(parts)-2]+olt[len(olt)-3:])
			fbs_cnt++
		}

		// 安装位置，二级分光
		cur, _ = source.GetCellValue("号线资源表", "B"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "B"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), cur+strconv.Itoa(sbs_cnt)+"号分光器")

		// 设备名称，二级分光
		olt, _ := source.GetCellValue("号线资源表", "P"+strconv.Itoa(i+3))
		parts := strings.Split(olt, "-")
		installation.SetCellValue("data sheet", "C"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), cur+"OBD0"+parts[len(parts)-2]+olt[len(olt)-3:]+"-0"+strconv.Itoa((sbs_cnt-1)%8+1))
		sbs_cnt++
	}

	// 设备编码 TODO:

	// 固定，设备型号
	for i := 0; i < FRIST_BEAM_SPLITTER+SECOND_BEAM_SPLITTER; i++ {
		installation.SetCellValue("data sheet", "E"+strconv.Itoa(i+2), "光分波导1×8SC")
		installation.SetCellValue("data sheet", "F"+strconv.Itoa(i+2), "华为")

		// TOOD: can modify
		installation.SetCellValue("data sheet", "R"+strconv.Itoa(i+2), "仁寿")
		installation.SetCellValue("data sheet", "S"+strconv.Itoa(i+2), "仁寿IMS局向")
		installation.SetCellValue("data sheet", "T"+strconv.Itoa(i+2), "是")
		installation.SetCellValue("data sheet", "X"+strconv.Itoa(i+2), "自建")
		installation.SetCellValue("data sheet", "AA"+strconv.Itoa(i+2), "是")
	}

	// 分光
	for i := 0; i < FRIST_BEAM_SPLITTER; i++ {
		installation.SetCellValue("data sheet", "G"+strconv.Itoa(i+2), "8")
		installation.SetCellValue("data sheet", "H"+strconv.Itoa(i+2), "1")
	}
	for i := 0; i < SECOND_BEAM_SPLITTER; i++ {
		cur, _ := source.GetCellValue("号线资源表", "Q"+strconv.Itoa(i+3))
		b := []byte(cur)
		installation.SetCellValue("data sheet", "G"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), string(b[len(b)-1]))
		installation.SetCellValue("data sheet", "H"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), "2")
	}

	// 所属OLT设备
	olt, _ := source.GetCellValue("号线资源表", "P3")
	o := strings.Split(olt, " ")[0]
	for i := 0; i < FRIST_BEAM_SPLITTER+SECOND_BEAM_SPLITTER; i++ {
		installation.SetCellValue("data sheet", "I"+strconv.Itoa(i+2), o)
	}

	// 上联OBD设备 || 上联OBD端口
	for i := 0; i < FRIST_BEAM_SPLITTER; i++ {
		obd, _ := installation.GetCellValue("data sheet", "C"+strconv.Itoa(i+2))
		for j := 0; j < 8; j++ {
			if i*8+j >= SECOND_BEAM_SPLITTER {
				break
			}
			installation.SetCellValue("data sheet", "K"+strconv.Itoa(i*8+j+FRIST_BEAM_SPLITTER+2), obd)
			installation.SetCellValue("data sheet", "L"+strconv.Itoa(i*8+j+FRIST_BEAM_SPLITTER+2), "CD0"+strconv.Itoa(j+1))
		}
	}

	// 二维码
	pre_fbs = ""
	fbs_cnt = 1
	for i := 0; i < SECOND_BEAM_SPLITTER; i++ {
		// 设备名称，一级分光 || 安装位置，一级分光
		cur, _ := source.GetCellValue("号线资源表", "I"+strconv.Itoa(i+3))
		if pre_fbs != cur {
			pre_fbs = cur
			// TODO: 文字切割
			installation.SetCellValue("data sheet", "U"+strconv.Itoa(fbs_cnt+1), cur[len(cur)-67:len(cur)-3])
			fbs_cnt++
		}

		// 安装位置，二级分光
		cur, _ = source.GetCellValue("号线资源表", "C"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "U"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), cur)
	}

	// 经纬度
	for i := 0; i < FRIST_BEAM_SPLITTER; i++ {
		v, _ := source.GetCellValue("一级光交号线资源表", "N"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "V"+strconv.Itoa(i+2), v)
		v, _ = source.GetCellValue("一级光交号线资源表", "O"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "W"+strconv.Itoa(i+2), v)
	}
	for i := 0; i < SECOND_BEAM_SPLITTER; i++ {
		cur, _ := source.GetCellValue("号线资源表", "N"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "V"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), cur)
		cur, _ = source.GetCellValue("号线资源表", "O"+strconv.Itoa(i+3))
		installation.SetCellValue("data sheet", "W"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), cur)
	}

	// OBD端口
	// TODO: 所属设备编码，端口编码
	// ==================================================================================================
	obd_ports := excelize.NewFile()
	defer func() {
		if err := obd_ports.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	index, _ = obd_ports.NewSheet("data sheet")
	defer func() {
		if err := obd_ports.SaveAs("source/端口端口.xlsx"); err != nil {
			fmt.Println(err)
		}
	}()
	obd_ports.SetActiveSheet(index)
	obd_ports_title := [...]string{"所属设备名称**", "所属设备编码**", "端口序号**", "设备编码**", "端口业务", "空闲状态**", "用户号码/宽带账号", "所属区域**", "错误信息", "错误编码"}
	for idx, val := range obd_ports_title {
		obd_ports.SetCellValue("data sheet", title_no[idx]+"1", val)
	}

	// obd sum
	OBD_SUM := FRIST_BEAM_SPLITTER * 8
	for i := 0; i < SECOND_BEAM_SPLITTER; i++ {
		cur, _ := source.GetCellValue("号线资源表", "Q"+strconv.Itoa(i+3))
		b := []byte(cur)
		installation.SetCellValue("data sheet", "G"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), string(b[len(b)-1]))
		installation.SetCellValue("data sheet", "H"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2), "2")

		v, _ := installation.GetCellValue("data sheet", "G"+strconv.Itoa(i+FRIST_BEAM_SPLITTER+2))
		n, _ := strconv.Atoi(v)
		OBD_SUM += n
	}

	// "端口业务", "空闲状态**", "所属区域**"
	for i := 0; i < OBD_SUM; i++ {
		obd_ports.SetCellValue("data sheet", "E"+strconv.Itoa(i+2), "FTTH")
		obd_ports.SetCellValue("data sheet", "F"+strconv.Itoa(i+2), "空闲")
		obd_ports.SetCellValue("data sheet", "H"+strconv.Itoa(i+2), "仁寿") // TODO:modify
	}

	// 所属设备名称，端口序号
	idx := 0
	for i := 0; i < FRIST_BEAM_SPLITTER+SECOND_BEAM_SPLITTER; i++ {
		v, _ := installation.GetCellValue("data sheet", "C"+strconv.Itoa(i+2))
		v1, _ := installation.GetCellValue("data sheet", "G"+strconv.Itoa(i+2))
		n, _ := strconv.Atoi(v1)
		for j := 0; j < n; j++ {
			obd_ports.SetCellValue("data sheet", "A"+strconv.Itoa(i+2+idx), v)
			obd_ports.SetCellValue("data sheet", "C"+strconv.Itoa(i+2+idx), "CD0"+strconv.Itoa(j+1))
			idx++
		}
		idx--
	}

	// 标准地址
	// ==================================================================================================
	address := excelize.NewFile()
	defer func() {
		if err := address.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	index, _ = address.NewSheet("data sheet")
	defer func() {
		if err := address.SaveAs("source/标准地址.xlsx"); err != nil {
			fmt.Println(err)
		}
	}()

	address.SetActiveSheet(index)
	address_title := [...]string{"七级地址ID**", "一级地址", "二级地址", "三级地址", "四级地址", "五级地址", "六级地址", "七级地址", "八级地址", "九级地址", "十级地址", "十一级地址", "设备类型**", "设备名称**", "接入业务**", "接入方式**", "已安装号码", "地址模式**", "地域类型", "区域类型", "子区域类型(7级及以上必填)", "所属楼层(10级及以上必填)", "地址归属(7级地址必填)", "地址归属(名单制楼宇)", "产权归属*", "覆盖区域*", "小区分类*(集团)", "小区类型*(集团)", "错误编码", "错误信息"}
	for idx, val := range address_title {
		address.SetCellValue("data sheet", title_no[idx]+"1", val)
	}
	ADDR := "data sheet"
	SOURCE_ADDR := "驻地网地址"

	address_no := 2
	for {
		idx, _ := excelize.CoordinatesToCellName(2, address_no)
		tmp, _ := source.GetCellValue("驻地网地址", idx)
		if isBlank(tmp) {
			address_no--
			break
		}
		address_no++
	}

	pre_e, cur := "", ""
	add_cnt := FRIST_BEAM_SPLITTER + 2
	for i := 2; i < address_no; i++ {
		a, _ := source.GetCellValue(SOURCE_ADDR, "B"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "A"+strconv.Itoa(i), a)
		b, _ := source.GetCellValue(SOURCE_ADDR, "C"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "B"+strconv.Itoa(i), b)
		c, _ := source.GetCellValue(SOURCE_ADDR, "D"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "C"+strconv.Itoa(i), c)
		d, _ := source.GetCellValue(SOURCE_ADDR, "F"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "D"+strconv.Itoa(i), d)
		e, _ := source.GetCellValue(SOURCE_ADDR, "G"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "E"+strconv.Itoa(i), e)
		f, _ := source.GetCellValue(SOURCE_ADDR, "H"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "F"+strconv.Itoa(i), f)
		g, _ := source.GetCellValue(SOURCE_ADDR, "I"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "G"+strconv.Itoa(i), g)
		h, _ := source.GetCellValue(SOURCE_ADDR, "J"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "H"+strconv.Itoa(i), h)
		ii, _ := source.GetCellValue(SOURCE_ADDR, "K"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "I"+strconv.Itoa(i), ii)
		j, _ := source.GetCellValue(SOURCE_ADDR, "L"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "J"+strconv.Itoa(i), j)
		k, _ := source.GetCellValue(SOURCE_ADDR, "M"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "K"+strconv.Itoa(i), k)
		l, _ := source.GetCellValue(SOURCE_ADDR, "n"+strconv.Itoa(i+1))
		address.SetCellValue(ADDR, "L"+strconv.Itoa(i), l)
		address.SetCellValue(ADDR, "M"+strconv.Itoa(i), "分光器")

		tmp_e, _ := source.GetCellValue(SOURCE_ADDR, "E"+strconv.Itoa(i+1))
		if tmp_e != pre_e {
			pre_e = tmp_e
			cur, _ = installation.GetCellValue("data sheet", "C"+strconv.Itoa(add_cnt))
			cur = cur[len(cur)-11:]
			add_cnt++
		}
		address.SetCellValue(ADDR, "N"+strconv.Itoa(i), c+d+e+f+g+h+ii+j+cur)

		address.SetCellValue(ADDR, "O"+strconv.Itoa(i), "宽带")
		address.SetCellValue(ADDR, "P"+strconv.Itoa(i), "FTTH")
		address.SetCellValue(ADDR, "R"+strconv.Itoa(i), "非到户")
		address.SetCellValue(ADDR, "S"+strconv.Itoa(i), "乡镇")
		address.SetCellValue(ADDR, "T"+strconv.Itoa(i), "城市小区")
		address.SetCellValue(ADDR, "U"+strconv.Itoa(i), "小区")
		address.SetCellValue(ADDR, "V"+strconv.Itoa(i), k[:len(k)-3])
		address.SetCellValue(ADDR, "W"+strconv.Itoa(i), "公众")
		address.SetCellValue(ADDR, "X"+strconv.Itoa(i), "公众地址")
		address.SetCellValue(ADDR, "Y"+strconv.Itoa(i), "自有")
		address.SetCellValue(ADDR, "Z"+strconv.Itoa(i), "城市")
		address.SetCellValue(ADDR, "AA"+strconv.Itoa(i), "社区")
		address.SetCellValue(ADDR, "AB"+strconv.Itoa(i), "封闭式住宅小区")
	}

	// 小区维度
	// ==================================================================================================
	hood := excelize.NewFile()
	defer func() {
		if err := hood.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	index, _ = hood.NewSheet("data sheet")
	defer func() {
		if err := hood.SaveAs("source/小区维度.xlsx"); err != nil {
			fmt.Println(err)
		}
	}()

	hood.SetActiveSheet(index)
	hood_title := [...]string{"本地网**", "服务区**", "网格**", "小区名称**", "小区地址", "设备名称**", "错误编码", "错误信息"}
	for idx, val := range hood_title {
		hood.SetCellValue("data sheet", title_no[idx]+"1", val)
	}
	hood_no := SECOND_BEAM_SPLITTER

	hood_name, _ := source.GetCellValue(SOURCE_ADDR, "J"+strconv.Itoa(2))
	net, _ := source.GetCellValue(SOURCE_ADDR, "G"+strconv.Itoa(2))
	for i := 0; i < hood_no; i++ {
		hood.SetCellValue(ADDR, "A"+strconv.Itoa(i+2), "眉山") // TODO:
		hood.SetCellValue(ADDR, "B"+strconv.Itoa(i+2), "仁寿") // TODO:
		hood.SetCellValue(ADDR, "C"+strconv.Itoa(i+2), "仁寿"+net+"眉山网格")
		hood.SetCellValue(ADDR, "D"+strconv.Itoa(i+2), hood_name)

		hood_addr, _ := installation.GetCellValue("data sheet", "B"+strconv.Itoa(i+2+FRIST_BEAM_SPLITTER))
		hood.SetCellValue(ADDR, "E"+strconv.Itoa(i+2), hood_addr)
		install_name, _ := installation.GetCellValue("data sheet", "C"+strconv.Itoa(i+2+FRIST_BEAM_SPLITTER))
		hood.SetCellValue(ADDR, "F"+strconv.Itoa(i+2), install_name)
	}
}

func isBlank(s string) bool {
	// 移除字符串首尾的空白字符
	trimmed := strings.TrimSpace(s)

	// 判断是否为空字符串
	return trimmed == ""
}
