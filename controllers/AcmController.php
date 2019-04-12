<?php

namespace app\controllers;

use Yii;
use yii\filters\AccessControl;
use yii\web\Controller;
use yii\web\Response;
use yii\filters\VerbFilter;
use app\models\LoginForm;
use app\models\ContactForm;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class AcmController extends Controller {

    
    public $begin_day = 0;
    public $end_day = 0;
    public $all_date = [];
    
    private $event_arr = [
        'in'=>'Вход',
        'out'=>'Виход'
    ];
    /**
     * {@inheritdoc}
     */
    public function behaviors() {
        return [
            'access' => [
                'class' => AccessControl::className(),
                'only' => ['logout'],
                'rules' => [
                    [
                        'actions' => ['logout'],
                        'allow' => true,
                        'roles' => ['@'],
                    ],
                ],
            ],
            'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'logout' => ['post'],
                ],
            ],
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function actions() {
        return [
            'error' => [
                'class' => 'yii\web\ErrorAction',
            ],
            'captcha' => [
                'class' => 'yii\captcha\CaptchaAction',
                'fixedVerifyCode' => YII_ENV_TEST ? 'testme' : null,
            ],
        ];
    }

    /**
     * Displays homepage.
     *
     * @return string
     */
    public function actionIndex() {
//        set_time_limit(500); //
//        ini_set('max_execution_time', 900);
        $mem_start = memory_get_usage();
        $model = new \app\models\UploadForm();
        
        if(\Yii::$app->request->isPost && $model->load($_POST)){
            $model->excelFile = \yii\web\UploadedFile::getInstance($model, 'excelFile');
            if($model->upload()){
                $file = $dir = Yii::getAlias('@app/web/uploads/').$model->excelFile->name;
                
//                $inputFileType = ucfirst(strtolower($model->excelFile->extension));
                $inputFileName = $file;
                $inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($inputFileName);
                $reader =  \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
                $reader->setInputEncoding('CP1251');
                $reader->setReadDataOnly(true);
                $spreadsheet = $reader->load($inputFileName);
                $sheet = $spreadsheet->getActiveSheet();
                
                $confCell= [];
                $temp_id = [];
                $temp_date = [];
                
                $alldata = [];
                
                $users_array =[];
                
                $access_arr = ['Фенікс Агро'];
                
                $nColumn = \PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn());
                
                for($j = 1; $j <=$nColumn; $j++){
                    $head = $sheet->getCellByColumnAndRow($j, 1)->getValue();
                    switch ($head):
                    case 'Время':
                        $confCell['datetime'] = $j;
                        break;
                    case 'Номер карты':
                        $confCell['id_user'] = $j;
                        break;
                    case 'Контроллер':
                        $confCell['door'] = $j;
                        break;
                    case 'Описание':
                        $confCell['description'] = $j;
                        break;
                    case 'Имя':
                        $confCell['name'] = $j;
                        break;
                    case 'Фамилия':
                        $confCell['secondname'] = $j;
                        break;
                    case 'Отдел':
                        $confCell['access'] = $j;
                        break;
                    endswitch;
                }
                
                $user_data=[];
                
                for($i = 2;$i<= $sheet->getHighestRow();$i++){
                    $id_user = $sheet->getCellByColumnAndRow($confCell['id_user'], $i)->getValue();
                    $name_user = $sheet->getCellByColumnAndRow($confCell['name'], $i)->getValue();
                    $secondname_user = $sheet->getCellByColumnAndRow($confCell['secondname'], $i)->getValue();
                    $access_user = $sheet->getCellByColumnAndRow($confCell['access'], $i)->getValue();
                    if(!in_array($access_user, $access_arr)){
                        continue;
                    }
                    $datetime =  $sheet->getCellByColumnAndRow($confCell['datetime'], $i)->getValue();
                    $door =  $sheet->getCellByColumnAndRow($confCell['door'], $i)->getValue();
                    $description =  $sheet->getCellByColumnAndRow($confCell['description'], $i)->getValue();
                    
                    $date = date('d-m-Y', strtotime($datetime));
                    $time = date('H:i', strtotime($datetime));
                    if(!$this->begin_day){
                        $this->begin_day = $date;
                    }
                    if($this->begin_day> $date){
                        $this->begin_day = $date;
                    }
                    if($this->end_day!= $date){
                        $this->end_day = $date;
                    }
                    if(!in_array($date, $this->all_date)){
                        $this->all_date[] = $date;
                    }
                    if(!array_key_exists('date', $temp_date)){
                        $temp_date['date'] = $date;
                        $temp_date['time'][] = 
                        [$time, $this->getInOn($door, $description)];
                    }else if(array_key_exists('date', $temp_date)&&$temp_date['date'] ==$date){
                        $temp_date['time'][] =
                        [$time, $this->getInOn($door, $description)];
                    }else if(array_key_exists('date', $temp_date)&&$temp_date['date'] !=$date){

                        $user_data['date'][$temp_date['date']] = $this->calcTime($temp_date['time']);
                        
                        
                        $calc_time = $this->calcTime($temp_date['time'],$date);
                        
                        $temp_date = [];
                        $temp_date['date'] = $date;
                        $temp_date['time'][] =
                        [$time, $this->getInOn($door, $description)];
                    }
                    
                    if(!array_key_exists('id', $user_data)){
                        $user_data['id'] = $id_user;
                        $name = $sheet->getCellByColumnAndRow($confCell['name'], $i)->getValue();
                        $secondname = $sheet->getCellByColumnAndRow($confCell['secondname'], $i)->getValue();
                        $user_data['name'] = $name;
                        $user_data['secondname'] = $secondname;
                    }else if(array_key_exists('id', $user_data)&&$user_data['name']!=$name_user&&$user_data['secondname']!=$secondname_user){
                        $users_array[] = $user_data;
                        
                        $user_data = [];
                        $temp_date = [];
                        
                        $user_data['id'] = $id_user;
                        $name = $sheet->getCellByColumnAndRow($confCell['name'], $i)->getValue();
                        $secondname = $sheet->getCellByColumnAndRow($confCell['secondname'], $i)->getValue();
                        $user_data['name'] = $name;
                        $user_data['secondname'] = $secondname;
                    }
                }
                

                sort($this->all_date);
                
                $this->genereteExcel($users_array);
                die();
                
            }
        }
        return $this->render('index',['model'=>$model]);
    }
    private function calcTime($arr_time,$date = null){
        
        $max_time = '17:30';
        
        
        $first = null;
        $last = null;
        $all_min = 0;
        $all_not_work = 0;
        
        $last_in = null;
        
        $count = count($arr_time);
        if($count==0)
        {
            return null;
        }
        
        $last_event = null;
        $temp_time = [];
        for($i= 0;$i < count($arr_time);$i++){
            switch ($arr_time[$i][1]){
            case 'in':
                if($last_event==null&&$first==null){
                    $last_event = 'in';
                    $first = $arr_time[$i][0];
                }else{
                    
                }
                $temp_time[]['in'] = $arr_time[$i][0];
                $last_in = $arr_time[$i][0];
                break;
            case 'out':
                if($last_event=='in'){
                    end($temp_time);
                    $temp_time[key($temp_time)]['out'] = $arr_time[$i][0];
                }else{
                    $temp_time[]['out'] = $arr_time[$i][0];
                }
                $last = $arr_time[$i][0];
                break;
            case 'inOut':
                break;
            }
        }
        if($last_in>$last){
            if($last< $max_time){
                $last = $max_time;
                end($temp_time);
                $temp_time[key($temp_time)]['out'] = $last;
            }
        }
        $min_all = 0;
        $time = $this->getWorkAndOut($temp_time,$date);
        $temp_time['work'] = $time['work'];
        $temp_time['later_in'] = $time['later_in'];
        $temp_time['out'] = $time['out'];
        $temp_time['before_out'] = $time['before_out'];
        $temp_time['before_lunch'] = $time['before_lunch'];
        $temp_time['lunch'] = $time['lunch'];
        $temp_time['after_lunch'] = $time['after_lunch'];
        $temp_time['first'] = $first;
        $temp_time['last'] = $last;
        return $temp_time;
    }
    
    private function getInOn($door,$event){
        $arr_event = [
            'Турникет 1'=>[
                'Успешный вход'=>'out',
                'Успешный выход'=>'in'
            ],
            'Турникет 2'=>[
                'Успешный вход'=>'in',
                'Успешный выход'=>'out'
            ],
//            'Вихід двір'=>[
//                'Успешный вход'=>'inOut',
//                'Успешный выход'=>'inOut'
//            ]
        ];
        return $arr_event[$door][$event];
    }
    private function getWorkAndOut($arr_time,$date){
        
        $time_work = 0;
        $later_in = 0;
        $out_of_work = 0;
        $before_out = 0;
        
        $before_lunch = 0;
        $lunch = 0;
        $after_lunch = 0;
        
        $work_time =
            [
                [
                    '08:00',
                    '12:00'
                ],
                [
                    '12:00',
                    '13:00'
                ],
                [
                    '13:00',
                    '17:00'
                ]
            ];
        $weekday_time =
            [
                [
                    '08:00',
                    '19:00'
                ]
            ];
        
        $last_key = count($arr_time)-1;
        foreach ($arr_time as $key=>$value){
            if(isset($value['in'])){
                if($key==0){
                    $later_in = $this->TimeBeetwen($work_time[0][0],$value['in']);
                }
            }
            if(isset($value['in'])&&isset($value['out'])){
                $time_work+= $this->TimeBeetwen($value['in'],$value['out']);
            }
            if(isset($value['out'])){
                if($key==$last_key){
                    $before_out = $this->TimeBeetwen($value['out'],$work_time[2][1]);
                }
            }


        }
        $temp_last_out = 0;
        $temp_last_in = 0;
        $temp_last_event=false;
        
        foreach ($work_time as $key=>$value){
            
            $temp_last_out= 0;
            $temp_last_in= 0;
            foreach ($arr_time as $key_work=>$value_work){
                if($key=='0'){
                    if($key_work==0&&$value_work['in']>$value[1]){
                        $before_lunch = $this->TimeBeetwen($value[0],$value['1']);
                    }
                    if($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out==0){
                        $before_lunch = $this->TimeBeetwen($value[0],$value_work['in']);
                    }elseif($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out>0){
                        $before_lunch += $this->TimeBeetwen($temp_last_out,$value_work['in']);
                    }
                    if($value_work['out']>$value[0]&&$value_work['out']<$value[1]&&$temp_last_out==0){
                        $temp_last_out = $value_work['out'];
                    }
                }
                if($key=='1'){
                    if($temp_last_event=='in'){
                        
                    }
                    if($value_work['in']<$value[0]&&$value_work['out']>$value[1]){
//                        echo 'laung full<br/>';
                        continue;
                    }
                    if($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out==0){
                        $lunch = $this->TimeBeetwen($value[0],$value_work['in']);
                        $temp_last_in = $value_work['in'];
//                        echo $lunch.'='.$value[0].'+'.$value_work['in'].'<br/>';
                    }elseif($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out>0){
                        $lunch += $this->TimeBeetwen($temp_last_out,$value_work['in']);
                        $temp_last_in = $value_work['in'];
//                        echo $lunch.'='.$value[0].'+'.$value_work['in'].'<br/>';
                    }else{
//                        echo $value[0].'-'.$value_work['in'].'-'.$value[1].'<br/>';
                    }
                    if($value_work['out']>$value[0]&&$value_work['out']<$value[1]&&$temp_last_out==0){
                        $temp_last_out = $value_work['out'];
                        $lunch += $this->TimeBeetwen($value_work['out'],$value[1]);
                    }
                }
                if($key=='2'){
                    if($temp_last_event=='in'){
                        
                    }
                    if($value_work['in']<$value[0]&&$value_work['out']>$value[1]){
//                        echo 'laung full<br/>';
                        continue;
                    }
                    if($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out==0){
                        $after_lunch = $this->TimeBeetwen($value[0],$value_work['in']);
                        $temp_last_in = $value_work['in'];
//                        echo $after_lunch.'='.$value[0].'+'.$value_work['in'].'<br/>';
                    }elseif($value_work['in']>$value[0]&&$value_work['in']<$value[1]&&$temp_last_out>0){
                        $after_lunch += $this->TimeBeetwen($temp_last_out,$value_work['in']);
                        $temp_last_in = $value_work['in'];
//                        echo $after_lunch.'='.$value[0].'+'.$value_work['in'].'<br/>';
                    }else{
//                        echo $value[0].'-'.$value_work['in'].'-'.$value[1].'<br/>';
                    }
                    if($value_work['out']>$value[0]&&$value_work['out']<$value[1]&&$temp_last_out==0){
                        $temp_last_out = $value_work['out'];
                        $after_lunch += $this->TimeBeetwen($value_work['out'],$value[1]);
                    }
                }
            }
        }
            return array(
                'work'=>$this->TimeToStr($time_work),
                'later_in'=>$this->TimeToStr($later_in),
                'out'=>$this->TimeToStr($out_of_work),
                'before_out'=>$this->TimeToStr($before_out),
                
                'before_lunch'=>$before_lunch,
                'lunch'=>$lunch,
                'after_lunch'=>$after_lunch,
                );
    }
//    private function getOutOfWork($arr_time){
//        
//        $out_of_work = [];
//        $time_work = [];
//        
//        $work_time =
//            [
//                [
//                    '08:00',
//                    '12:00'
//                ],
//                [
//                    '12:00',
//                    '13:00'
//                ],
//                [
//                    '13:00',
//                    '17:00'
//                ]
//            ];
//        $weekday_time =
//            [
//                [
//                    '08:00',
//                    '19:00'
//                ]
//            ];
//        
//        $i = 0;
//            //робочий диапазон
//            foreach ($work_time as $work_range){
//                $end_work_range = false;
//                $time_work[$i] = 0;
//                $out_of_work[$i] = 0;
//                //посещаемость человека
//                for($j=0;$j< count($arr_time);$j++){
//                    //если первый период
//                    if($i==0){
//                        //если человек опоздал
//                        if($work_range[0]<$arr_time[$j]['in']){
//                            //если человек не пришел до конца рабочего периода
//                            if($work_range[1]<$arr_time[$j]['in']){
//                                $time_work[0] = 0;
//                                $out_of_work[0] += $this->TimeBeetwen($work_range[0],$work_range[1]);
//                                $end_work_range = true;
//                            }else {//если человек пришел до конца рабочего периода
//                                // и если человек ушел позже конца рабочего периода
//                                if($work_range[1]<$arr_time[$j]['out']){
//                                    $time_work[0] = $this->TimeBeetwen($arr_time[$j]['in'],$work_range[1]);
//                                    $out_of_work[0] += $this->TimeBeetwen($work_range[0],$arr_time[$j]['in']);
//                                    $end_work_range = true;
//                                }else{// и если человек ушел раньше конца рабочего периода
//                                    $time_work[0] += $this->TimeBeetwen($arr_time[$j]['in'],$work_range[1]);
//                                    //но вернулся до конца рабочего времени периода
//                                    if(key_exists($j+1, $arr_time)&&$arr_time[$j+1]['in']<$work_range[1]){
//                                        $out_of_work[0] += $this->TimeBeetwen($work_range[0],$arr_time[$j]['in']) + $this->TimeBeetwen($arr_time[$j]['out'],$arr_time[$j+1]['in']);
//                                    }else{
//                                        //если не вернулся до конца рабочего периода
//                                        $out_of_work[0] += $this->TimeBeetwen($work_range[0],$arr_time[$j]['in']) + $this->TimeBeetwen($arr_time[$j]['out'],$work_range[1]);
//                                        $end_work_range = true;
//                                    }
//                                }
//                            }
//                        }else{
//                            //если человек ушел позже конца рабочего периода
//                             if($work_range[1]<=$arr_time[$j]['out']){
//                                $time_work[0] = $this->TimeBeetwen($arr_time[$j]['in'],$work_range[1]);
//                                $out_of_work[0] = 0;
//                            }else {//если человек ушел до конца рабочего периода
//                                $time_work[0] += $this->TimeBeetwen($arr_time[$j]['in'],$work_range[1]);
////                                $out_of_work[0] = $this->TimeBeetwen($arr_time[$j]['out'],$work_range[1]);
//                                 if(key_exists($j+1, $arr_time)&&$arr_time[$j+1]['in']<$work_range[1]){
//                                        $out_of_work[0] += $this->TimeBeetwen($work_range[0],$arr_time[$j]['in']) + $this->TimeBeetwen($arr_time[$j]['out'],$arr_time[$j+1]['in']);
//                                    }else{
//                                        //если не вернулся до конца рабочего периода
//                                        $out_of_work[0] += $this->TimeBeetwen($work_range[0],$arr_time[$j]['in']) + $this->TimeBeetwen($arr_time[$j]['out'],$work_range[1]);
//                                        $end_work_range = true;
//                                    }
//                            }
//                        }
//                    }else if($work_range[1]>$arr_time[$j]['in'] &&$end_work_range==false){
//                        continue;
//                        
//                    }else if($work_range[1]>$arr_time[$j]['in'] &&$j=count($arr_time)){
//                        
//                    }
//                    if($end_work_range){
//                        $hour = floor($out_of_work[$i]/60);
//                        $min = $out_of_work[$i]%60;
//                        $out_of_work[$i] = $hour.':'.$min;
//                        $i++;
//                        continue 2;
//                    }
////                    if($work_range[0]>$range['in'] && $work_range[0]<$range['out']){
////                        
////                        if($range['in']> $work_range[0]){
////    //                    echo $work_range[0].'-'.$range['in'].'<br/>';
////                            $out_of_work[$i]+= $this->TimeBeetwen($work_range[0],$range['in']);
////                        }
////                        if($range['out'] < $work_range[1]){
////
////                            $out_of_work[$i] += $this->TimeBeetwen($range['out'],$work_range[1]);
////                        }
////                        $hour = floor($out_of_work[$i]/60);
////                        $min = $out_of_work[$i]%60;
////                        $out_of_work[$i] = $hour.':'.$min;
////                        $i++;
////                        continue 2;
////                    }else if($range['out']<$work_range[0]){
////                        $i++;
////                        continue 2;
////                    }
//                }
////                    if($range['in']<$work_range[1]){
////                        
////                        if($range['in']> $work_range[0]){
////    //                    echo $work_range[0].'-'.$range['in'].'<br/>';
////                            $out_of_work[$i]+= $this->TimeBeetwen($work_range[0],$range['in']);
////                        }
////                        if($range['out'] < $work_range[1]){
////
////                            $out_of_work[$i] += $this->TimeBeetwen($range['out'],$work_range[1]);
////                        }
////                        $hour = floor($out_of_work[$i]/60);
////                        $min = $out_of_work[$i]%60;
////                        $out_of_work[$i] = $hour.':'.$min;
////                        $i++;
////                        continue 2;
////                    }else if($range['out']<$work_range[0]){
////                        $i++;
////                        continue 2;
////                    }
////                $i++;
//            }
//        return $out_of_work;
//    }

    private function TimeBeetwen($min,$max){
        if($max <= $min){
            return 0;
        }
        $beetwen = (strtotime($max) - strtotime($min))/60;
        return $beetwen;
    }
    private function TimeToStr($beetwen){
        $hour = floor($beetwen/60);
        $min = $beetwen%60;
        return ($hour>9?$hour:'0'.$hour).':'.($min>9?$min:'0'.$min);
    }
    private function genereteExcel($date){
        
        \PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Calc info');
        
        $sheet1 = $spreadsheet->createSheet();
        $sheet1->setTitle('Detail');
        

        $sheet->getColumnDimension('A')->setWidth(12);
        $sheet->getColumnDimension('B')->setWidth(25);
        $sheet->getColumnDimension('C')->setAutoSize(true);
        $sheet->getColumnDimension('E')->setAutoSize(true);
        $sheet->getColumnDimension('F')->setAutoSize(true);
        $sheet->getColumnDimension('H')->setAutoSize(true);
        $sheet->getColumnDimension('I')->setAutoSize(true);
        $sheet->getColumnDimension('J')->setAutoSize(true);
        $sheet->getColumnDimension('G')->setAutoSize(true);
        
        $sheet1->getColumnDimension('A')->setAutoSize(true);
        $sheet1->getColumnDimension('B')->setAutoSize(true);
        $sheet1->getColumnDimension('C')->setAutoSize(true);
        $sheet1->getColumnDimension('E')->setAutoSize(true);
        $sheet1->getColumnDimension('F')->setAutoSize(true);
        $sheet1->getColumnDimension('H')->setAutoSize(true);
        $sheet1->getColumnDimension('I')->setAutoSize(true);
        $sheet1->getColumnDimension('J')->setAutoSize(true);
        $sheet1->getColumnDimension('G')->setAutoSize(true);
        
        
        $sheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 1);
        $sheet->freezePane('A2');
        $start_num = 3;
        $row_num = $start_num;
        
        $sheet->setCellValue('A1', 'Номер карти');
        $sheet->setCellValue('B1', 'ФІО');
        $sheet->setCellValue('C1', 'Дата');
        $sheet->setCellValue('D1', 'Прихід');
        $sheet->setCellValue('E1', 'Запізнення');
        $sheet->setCellValue('F1', 'Вихід');
        $sheet->setCellValue('G1', 'Ранній вихід');
        $sheet->setCellValue('H1', 'Пізній вихід');
        $sheet->setCellValue('I1', 'Часи на роботі');
        $sheet->setCellValue('J1', 'Відсутність');
        $sheet->setCellValue('J2', 'пізніше першого входу');
        $sheet->setCellValue('J3', 'і раніше виходу');
        $sheet->setCellValue('K1', 'До обіду');
        $sheet->setCellValue('L1', 'Обід(1 час)');
        $sheet->setCellValue('M1', 'Після обіду');
        
        $styleSumArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                    'color' => ['argb' => 'FFFF0000'],
                ],
            ],
            'alignment'=>[
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ]
        ];
        
        $sheet->setCellValue('E'.$row_num, '08:00');
//        $sheet->setCellValue('F1', 'Вихід');
        $sheet->setCellValue('G'.$row_num, '17:00');
        $sheet->setCellValue('I'.$row_num, '=G'.$row_num.'-E'.$row_num);
//        $sheet->setCellValue('I1', 'Обід(1 час)');
//        $sheet->setCellValue('A1', '');
        
        $id = NULL;
        
        foreach ($date as $key=>$data){
            $row = $row_num+1;
            $row_end = $row_num + (count($data['date'])>0?count($data['date']):1);
            $row_num = $row_end;
            
            $sheet->mergeCellsByColumnAndRow(1,$row,1,$row_end);
            $sheet->setCellValue('A'.$row, $data['id']);
            
            $sheet->mergeCells('B'.$row.':B'.$row_end);
            $sheet->setCellValue('B'.$row, $data['name'].' '.$data['secondname']);
            
            $j=$row;
            
            if($id==null){
                $id = $data['id'];
            }
            if(count($data['date'])>0){
                foreach($data['date'] as $key=>$value){

                    $sheet->setCellValue('C'.$j, $key);
                    if($value['first']!=''){
                        $sheet->setCellValue('D'.$j, $value['first']);
                        $sheet->setCellValue('E'.$j, '=IF(E'.$start_num.'<D'.$j.',(D'.$j.'-E'.$start_num.')*60*24,0)');
                        $sheet->setCellValue('F'.$j, $value['last']);
                        $sheet->setCellValue('G'.$j, '=IF(G'.$start_num.'>F'.$j.',(G'.$start_num.'-F'.$j.')*60*24,0)');
                        $sheet->setCellValue('H'.$j, '=IF(G'.$start_num.'<F'.$j.',(F'.$j.'-G'.$start_num.')*60*24,)');
                        $sheet->setCellValue('I'.$j, $value['work']);
                        $sheet->setCellValue('J'.$j, '=IF(AND(I'.$j.'>0,I'.$start_num.'>I'.$j.'),(I'.$start_num.'-I'.$j.')*60*24,0)');
                        $sheet->setCellValue('K'.$j, $value['before_lunch']);
                        $sheet->setCellValue('L'.$j, $value['lunch']);
                        $sheet->setCellValue('M'.$j, $value['after_lunch']);
                    }else{
                        $sheet->setCellValue('D'.$j, '-');
                        $sheet->setCellValue('E'.$j, '-');
                    }


                    $j++;
                }
                $sheet->setCellValue('E'.($row_end+1), '=SUM(E'.$row.':E'.($row_end).')');
                $sheet->setCellValue('G'.($row_end+1), '=SUM(G'.$row.':G'.($row_end).')');
                $sheet->setCellValue('H'.($row_end+1), '=SUM(H'.$row.':H'.($row_end).')');
                $sheet->setCellValue('I'.($row_end+1), '=SUM(I'.$row.':I'.($row_end).')');
                $sheet->setCellValue('J'.($row_end+1), '=SUM(J'.$row.':J'.($row_end).')');
                $sheet->setCellValue('K'.($row_end+1), '=SUM(K'.$row.':K'.($row_end).')');
                $sheet->setCellValue('L'.($row_end+1), '=SUM(L'.$row.':L'.($row_end).')');
                $sheet->setCellValue('M'.($row_end+1), '=SUM(M'.$row.':M'.($row_end).')');
                
                $sheet->mergeCells('A'.($row_end+1).':B'.($row_end+1));
                $sheet->setCellValue('A'.($row_end+1), $data['secondname'].' '.$data['name']);
                
                $sheet->getStyle('A'.($row_end+1).':M'.($row_end+1))->applyFromArray($styleSumArray);
                
                $id = $data['id'];
                $row_num++;
                $j++;
            }
            
        }

        $id = NULL;
        $row_num = 3;
//        echo $j.'asd';
        foreach ($date as $key=>$data){

            $j1=$row_num;
            
            if($id==null){
                $id = $data['id'];
            }
            $arr_exception = ['work','later_in','out','before_out','before_lunch','lunch','after_lunch','first','last'];
            if(count($data['date'])>0){
                foreach($data['date'] as $key_date=>$value_date){
                    if(count($value_date)>0){
                        foreach($value_date as $key_time=>$value_time){
                            if(in_array($key_time, $arr_exception,true)){
                                continue;
                            }
//                        if(count($value_date)>0){
                            foreach($value_time as $key_event=>$value_event){
                                $sheet1->setCellValue('A'.$j1, $data['id']);
                                $sheet1->setCellValue('B'.$j1, $data['name'].' '.$data['secondname']);
                                $sheet1->setCellValue('C'.$j1, $key_date);
//                                $sheet1->setCellValue('D'.$j, $key_time);
                                $sheet1->setCellValue('D'.$j1, $this->event_arr[$key_event]);
                                $sheet1->setCellValue('E'.$j1, $value_event);
                        $j1++;
                $row_num++;
                            }

//                        }
                        }
                    }
                }

                $id = $data['id'];
                $row_num++;
                $j1++;
            }
            
        }
        $sheet->getStyle('A'.$start_num.':I'.$j)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle('A1:L1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
        
        $sheet->getStyle('D2:D'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME3
        );
        $sheet->getStyle('F2:F'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME3
        );
        $sheet->getStyle('A'.$start_num.':K'.$start_num)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME3
        );
        $sheet->getStyle('E4:E'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
        );
        $sheet->getStyle('G4:G'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
        );
        $sheet->getStyle('H4:H'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
        );
        $sheet->getStyle('I4:I'.$j)->getNumberFormat()->setFormatCode('[HH]:MM:SS');
        $sheet->getStyle('J4:J'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
        );
        $sheet->getStyle('K4:K'.$j)->getNumberFormat()->setFormatCode(
            \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
        );
        
        $sheet->getStyle('B3:B'.($row_end+1))->getAlignment()->setWrapText(true);
        $sheet->setShowGridlines(true);
        
        $spreadsheet->getActiveSheet()->setAutoFilter('A1:B'.$j);
        
        
        $writer = new Xlsx($spreadsheet);
//        // redirect output to client browser
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="login_"'.$this->begin_day.'"_to_"'.$this->end_day.'".xlsx"');
        header('Cache-Control: max-age=0');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        
//        $writer->save("login_".$this->begin_day."_to_".$this->end_day.".xlsx");
    }
}
