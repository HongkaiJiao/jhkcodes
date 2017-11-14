<?php
namespace Home\Controller;
use Think\Controller;
class ExelController extends Controller {
    public function index(){
        $this->display();
    }
    /**
     *
     * 导出Excel方法
     */
    public function exportExcel($expTitle,$expCellName,$expTableData){
        $xlsTitle = iconv('utf-8', 'gb2312', $expTitle);//文件名称
        $fileName = $expTitle.date('_YmdHis');//or $xlsTitle 文件名称可根据自己情况设定
        $cellNum = count($expCellName);
        $dataNum = count($expTableData);

        vendor("PHPExcel.PHPExcel");//导入第三方类

        $objPHPExcel = new \PHPExcel();
        $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');
        //设置宽度
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);

        //设置行高
        $objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(30);//第一行
        $objPHPExcel->getActiveSheet()->getRowDimension('2')->setRowHeight(20);//第二行
        //设置字体样式
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10); //默认字体大小
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(16)->setBold(true);  //第一行 标题
        $objPHPExcel->getActiveSheet()->getStyle('A2:E2')->getFont()->setBold(true); //粗体    第二行
        //设置垂直、水平居中
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()
            ->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER)  //设置垂直
            ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);  //设置水平
        $objPHPExcel->getActiveSheet()->getStyle('A2:E2')->getAlignment()
            ->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER)   //设置垂直
            ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //设置水平

        //设置边框
        $objPHPExcel->getActiveSheet()->getStyle('A2:E2')->getBorders()->getAllBorders()
            ->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);

        $objPHPExcel->getActiveSheet(0)->mergeCells('A1:'.$cellName[$cellNum-1].'1');//合并单元格  mergeCells(a,b,c,d)合并单元格函数
//        常数	             值	    描述
//flexMergeNever	         0	    不显示。包含相同内容的单元不分组。这是缺省设置。
//flexMergeFree	             1	    自由。包含相同内容的单元总是合并。
//flexMergeRestrictRows	     2	    限制行。只有行中包含相同内容的相邻单元（向当前单元左边）才合并。
//flexMergeRestrictColumns	 3	    限制列。只有列中包含相同内容的相邻单元（向当前单元上方）才合并。
//flexMergeRestrictBoth	     4	    限制行和列。只有在行中（向左）或在列中（向上）包含相同内容的单元才合并。

        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', $expTitle.'  Export time:'.date('Y-m-d H:i:s'));  //表头

        for($i=0;$i<$cellNum;$i++){
            //-------------------sheet码--------------设置单元格的值----列号-----行号-----第几个array--0-第一个值-1-第二个值
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'2', $expCellName[$i][1]);
        }
        // Miscellaneous glyphs, UTF-8
        for($i=0;$i<$dataNum;$i++){
            for($j=0;$j<$cellNum;$j++){
                //-------------------sheet码--------------设置单元格的值----列号-----行号-----第几个array--0-第一个值-1-第二个值
                $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+3), $expTableData[$i][$expCellName[$j][0]]);
            }
        }

        header('pragma:public');
        header('Content-type:application/vnd.ms-excel;charset=utf-8;name="'.$xlsTitle.'.xls"');
        header("Content-Disposition:attachment;filename=$fileName.xls");//attachment新窗口打印inline本窗口打印
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }

    /**
     *
     * 导出Excel
     */
    function expUser(){//导出Excel
        $xlsName  = "Contacts";
        $xlsCell  = array(
            array('id','账号序列'),
            array('name','姓名'),
            array('tid','所属乡镇'),
            array('danwei','单位'),
            array('phone','电话')
        );
        $xlsModel = M('Contacts');
        $xlsData  = $xlsModel->Field('id,tid,name,danwei,phone')->select();
        $this->exportExcel($xlsName,$xlsCell,$xlsData);

    }
    /**
     *
     * 显示导入页面 ...
     */

    /**实现导入excel
     **/
    function impUser(){
        if (!empty($_FILES)) {
            $upload = new \Think\Upload();// 实例化上传类
            $filepath='./Public/Excle/';
            $upload->exts=array('xlsx','xls');// 设置附件上传类型   upload.class.php
            $upload->rootPath=$filepath; // 设置附件上传根目录
            $upload->saveName='time';    //上传文件命名
            $upload->autoSub=false;     //自动保存
            if (!$info=$upload->upload()) {   //upload（）上传文件方法
                $this->error($upload->getError());//检测上传
            }
//            print_r($info);exit;  $info=Array ( [import] => Array ( [name] => daorudaochu.xlsx [type] => application/octet-stream [size] => 8909 [key] => import [ext] => xlsx [md5] => a8ca390fffb70125f541ab410386c1e2 [sha1] => 7f1170c2ffd1780b7268986e9b04fa9c1d95066e [savename] => 1509175025.xlsx [savepath] => ) )
            foreach ($info as $key => $value) {
                unset($info);
                $info[0]=$value;
                $info[0]['savepath']=$filepath;
            }
            //print_r($info);exit;Array ( [0] => Array ( [name] => daorudaochu.xlsx [type] => application/octet-stream [size] => 8909 [key] => import [ext] => xlsx [md5] => a8ca390fffb70125f541ab410386c1e2 [sha1] => 7f1170c2ffd1780b7268986e9b04fa9c1d95066e [savename] => 1509330105.xlsx [savepath] => ./Public/Excle/ ) )

            vendor("PHPExcel.PHPExcel"); //导入第三方类

            $file_name=$info[0]['savepath'].$info[0]['savename'];
            $objReader = \PHPExcel_IOFactory::createReader('Excel5');         //通过已有的模板来创建空白文档
            $objPHPExcel = $objReader->load($file_name,$encode='utf-8');      //通过已有的模板来创建空白文档
            $sheet = $objPHPExcel->getSheet(0);   //读取
            $highestRow = $sheet->getHighestRow(); // 取得总行数
            $highestColumn = $sheet->getHighestColumn(); // 取得总列数

            $j=0;
            //echo "<pre>";print_r($highestRow);exit;
            for($i=3;$i<=$highestRow;$i++)
            {
                $data['name']= $objPHPExcel->getActiveSheet()->getCell("B".$i)->getValue();  //读取单元格
                $data['tid']= $objPHPExcel->getActiveSheet()->getCell("C".$i)->getValue();
                $data['danwei']= $objPHPExcel->getActiveSheet()->getCell("D".$i)->getValue();
                $data['phone']= $objPHPExcel->getActiveSheet()->getCell("E".$i)->getValue();
                M('contacts')->add($data);
                $j++;
            }

            unlink($file_name);  //删除文件
            $this->success('导入成功！本次导入联系人数量：'.$j);
        }else{
            $this->error("请选择上传的文件");
        }
    }
}