<?php


namespace App\Http\Controllers;


use App\Http\Controllers\Boxs\BoxTraits;
use App\Imports\ImportFamilyExcel;
use App\Imports\ImportNewExcelDataBase;
use App\Models\Farm;
use App\Models\ForestDataBase;
use App\Models\ScientificName;
use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Redirect;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\TblWidth;

class DataBaseController extends Controller
{
    use BoxTraits;

    public function __construct()
    {
        $this->middleware('auth');
    }

    public function importStore(Request $request)
    {
        // dd($request->get('reeplay'));
        if (ForestDataBase::where('farm_id',$request->get('farm_id'))->where('year',Carbon::now()->year)->count() && $request->get('reeplay') != 'on') {
            return redirect('/home')->with('status','Esta base de datos ya existe si desea reemplazar marque el check');
        }
        if (ForestDataBase::where('farm_id',$request->get('farm_id'))->where('year',Carbon::now()->year)->count() > 0 && $request->get('reeplay') == 'on' || $request->get('reeplay') == true) {
            DB::statement("DELETE FROM `forest_database` WHERE `farm_id` = ".$request->get('farm_id')." AND `year` = ".Carbon::now()->year);
        }
        $file=$request->file();
        $import=new ImportNewExcelDataBase($request->get('farm_id'));

        Excel::import($import,$file['data_base']);
        return redirect()->back();
    }

    public function clearFiles()
    {
        //File::delete(File::glob('public/*.docx'));

        $files = File::glob('*.docx'); //obtenemos todos los nombres de los ficheros

        foreach($files as $file){

            if(is_file($file))
                unlink($file); //elimino el fichero
        }
        return redirect()->back();
    }

    public function reportWord(Request $request)
    {
        try {
            $sysconf=Farm::find($request->get('farm_id'));
            Log::info(json_encode($sysconf));
            $phpWord=new PhpWord;
            // Establecer el estilo predeterminado
            $phpWord->setDefaultFontName('Times New Roma');                              //Fuente
            $phpWord->setDefaultFontSize(16);                                            //Tama??o de fuente
            $phpWord->setDefaultParagraphStyle(['vAlign'=>'center']);                    //Tama??o de fuente

            //A??adir p??gina
            $section=$phpWord->createSection(['orientation'=>'landscape']);

            // Agregar directorio
            $styleTOC=['tabLeader'=>\PhpOffice\PhpWord\Style\TOC::TABLEADER_DOT];
            $styleFont=['spaceAfter'=>60,'name'=>'Times New Roma','size'=>10];
            $section->addTOC($styleFont,$styleTOC);
            /**
             *
             * //Estilo por Defecto
             * $section->addText('Proyecto SIEPAC',['size'=>16,'bold'=>true]);
             * $section->addTextBreak();// Salto de l??nea
             * $section->addText('Inventario de Predios del Tramo 16',['size'=>16]);
             * $section->addTextBreak();// Salto de l??nea
             *
             *
             * $section->addText('Informe del Predio '.$sysconf->id_predio,['size'=>14,'bold'=>true]);
             * $section->addTextBreak();// Salto de l??nea
             * $section->addText($sysconf->owner,['size'=>14,'bold'=>true]);
             * $section->addTextBreak();// Salto de l??nea
             * $section->addTextBreak();// Salto de l??nea
             * $section->addTextBreak();// Salto de l??nea
             * $section->addText('Elaborado por:',['size'=>14]);
             * $section->addText('Ing. ??lvaro Sibaja Villegas',['size'=>14]);
             * $section->addTextBreak(7);// Salto de l??nea
             *
             *
             * $section->addText('San Jos??, Costa Rica ',['size'=>14]);
             *
             * $section->addTextBreak(5);// Salto de l??nea
             *
             * $section->addText('1. Consideraciones generales',['bold'=>true]);
             * $section->addTextBreak();// Salto de l??nea
             * $fontStyle=new \PhpOffice\PhpWord\Style\Font;
             * $fontStyle->setAuto(true);
             * $fontStyle->setName('Tahoma');
             * $fontStyle->setSize(13);
             * $myTextElement=$section->addText('Los datos base para la elaboraci??n del presente informe se obtuvieron
             * en primer lugar de la informaci??n proporcionada por el Ing. Allan Montoya, que comprende los shapefile de los predios, l??neas de tensi??n, torres y servidumbres; en segundo lugar, se necesit?? recolectar datos de campo para conocer la ubicaci??n y caracter??sticas de las especies dentro y fuera de la servidumbre.  En el inventario se ubicaron todos los ??rboles con un di??metro normal mayor a 15 cm, los cuales se georreferenciaron con un GPS Garmin 64s.');
             * $myTextElement->setFontStyle($fontStyle);
             *
             *
             * $section->addText('A continuaci??n, se presenta la nomenclatura utilizada en el presente informe.');
             *
             *
             * $nameScien = ForestDataBase::where('farm_id',$sysconf->id)->groupBy('name_cientifict')->count();
             * $section->addText('El inventario forestal registr??    '.$nameScien.'
             * especies diferentes, distribuidas en    10
             * familias. Un total de
             * 71    individuos fueron medidos, de estos    58     pertenecientes a especies con valor comercial,
             * mientras que    13    a especies sin valor comercial. Las ??reas basim??tricas (g) sumaron en total
             * 6.754    m2,  adem??s, se encontr?? un volumen comercial (Vc) de    6.414    m3 y el volumen total
             * (Vt) alcanz??    56.269    m3 (Cuadro 1).
             * ');*/
            $styleTable=[
                'borderSize'       =>1,
                'borderTopColor'   =>'1A446C',
                'borderRightColor' =>'FFFFFF',
                'borderLeftColor'  =>'FFFFFF',
                'borderBottomColor'=>'1A446C',
                'bold'             =>true,
                'cellMargin'       =>5,
                'size'             =>10,
                'width'            =>100,
                'vAlign'           =>'center',
            ];
            $styleFirstRow=[
                'borderSize'       =>1,
                'borderTopColor'   =>'1A446C',
                'borderRightColor' =>'FFFFFF',
                'borderLeftColor'  =>'FFFFFF',
                'borderBottomColor'=>'1A446C',
                'textDirection'    =>'tbRl',
                'vAlign'           =>'center'];
            // Estilo de primera l??nea
            $styleFirstColumn=[
                'borderSize'       =>1,
                'borderTopColor'   =>'FFFFFF',
                'borderRightColor' =>'FFFFFF',
                'borderLeftColor'  =>'FFFFFF',
                'borderBottomColor'=>'1A446C',
                'textDirection'    =>'tbRl',
                'vAlign'           =>'center'];
            // Estilo de primera l??nea

            /**
             * cuadro 1
             */
            $this->boxWordOne($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$styleFirstColumn);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 2
             */
            $this->boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 3
             */
            $this->boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 4
             */
            $this->boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 5
             */
            $this->boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 6
             */
            $this->boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 7
             */
            $this->boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 8
             */
            $this->boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea
            /**
             * cuadro 8
             */
            $this->boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
            $section->addTextBreak();// Salto de l??nea

            // El documento generado es Word2007
            $writer=\PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
            $writer->save('reporte_'.$sysconf->id_predio.'_uno.docx');
            return redirect()->back();
        }catch (\Exception $e){

            return  Redirect::back()->withErrors(['msg' => 'Revise los nombres Cient??ficos que tenga los nombres comunes agregado']);
        }

    }

    public function reportWordTwo(Request $request)
    {
        $sysconf=Farm::find($request->get('farm_id'));

        $phpWord=new PhpWord;
        // Establecer el estilo predeterminado
        $phpWord->setDefaultFontName('Times New Roma');                              //Fuente
        $phpWord->setDefaultFontSize(16);                                            //Tama??o de fuente
        $phpWord->setDefaultParagraphStyle(['vAlign'=>'center']);                    //Tama??o de fuente

        //A??adir p??gina
        $section=$phpWord->createSection(['orientation'=>'landscape']);

        // Agregar directorio
        $styleTOC=['tabLeader'=>\PhpOffice\PhpWord\Style\TOC::TABLEADER_DOT];
        $styleFont=['spaceAfter'=>60,'name'=>'Times New Roma','size'=>10];
        $section->addTOC($styleFont,$styleTOC);

        $styleTable=[
            'borderSize'       =>1,
            'borderTopColor'   =>'1A446C',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'bold'             =>true,
            'cellMargin'       =>5,
            'size'             =>10,
            'width'            =>100,
            'vAlign'           =>'center',
        ];
        $styleFirstRow=[
            'borderSize'       =>1,
            'borderTopColor'   =>'1A446C',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'vAlign'           =>'center'];
        // Estilo de primera l??nea
        $styleFirstColumn=[
            'borderSize'       =>1,
            'borderTopColor'   =>'FFFFFF',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'vAlign'           =>'center'];
        // Estilo de primera l??nea

        /**
         * cuadro 1
         */
        $this->boxWordOne($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$styleFirstColumn);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 2
         */
        $this->boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 3
         */
        $this->boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 4
         */
        $this->boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 5
         */
        $this->boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 6
         */
        $this->boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 7
         */
        $this->boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 8
         */
        $this->boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,0);
        $section->addTextBreak();// Salto de l??nea
/**
         * cuadro 9
         */
        $this->boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 10
         */
        $this->boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 11
         */
        $this->boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 12
         */
        $this->boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 13
         */
        $this->boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 14
         */
        $this->boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
        /**
         * cuadro 15
         */
        $this->boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de l??nea
         /**
         * cuadro 16
         */
        $this->boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de l??nea

        // El documento generado es Word2007
        $writer=\PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
        $writer->save('reporte_'.$sysconf->id_predio.'_dos.docx');

        return redirect()->back();
    }

    public function boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord)
    {
        $section->addText('Anexo 1. ??rboles inventariados en el Predio  '.$sysconf->id_predio);
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Anexo1',$styleTable,$styleFirstRow);


        $table=$section->addTable('Anexo1');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(500)->addText('Vano',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('??rbol',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Nombre Cient??fico',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Cobertura',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Comercial',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Servidumbre',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('??rea de Protecci??n',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Dap (cm)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Ht (m)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Hc (m)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('G (m??)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Vt (m??)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Vc (m??)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Coord X',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Coord Y',['bold'=>true,'size'=>9]);


        $datas=ForestDataBase::where('farm_id',$sysconf->id)->orderBy('name_cientifict','ASC')->get();

        foreach ($datas AS $data) {
            $table->addRow(5);// Altura de l??nea 400
            $table->addCell(500)->addText($data->vano,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->tree,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->family,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->name_cientifict,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->name_common,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->coverage,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->commercial,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->servitude,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->protection_area,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->dap,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->ht_m,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText($data->hc_m,['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText(round($data->g_m,3,PHP_ROUND_HALF_UP),['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText(round($data->vt_m,3,PHP_ROUND_HALF_UP),['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText(round($data->vc_m,3,PHP_ROUND_HALF_UP),['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText(round($data->coord_x,3,PHP_ROUND_HALF_UP),['bold'=>true,'size'=>9]);
            $table->addCell(500)->addText(round($data->coord_y,3,PHP_ROUND_HALF_UP),['bold'=>true,'size'=>9]);
        }

    }

    public function boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero = 0)
    {
        if ($report && $numero==0) {
            $section->addText('Cuadro 8. Distribuci??n diam??trica (cm) del volumen comercial (m3) de ??rboles ubicados en ??rea de Protecci??n del Predio '.$sysconf->id_predio);
        } elseif ($report && $numero==0) {
            $section->addText('Cuadro 8. Distribuci??n diam??trica (cm) del volumen comercial (m3) de ??rboles en ??reas de Protecci??n dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }elseif( $numero==3) {
            $section->addText('Cuadro 15. Distribuci??n diam??trica (cm) del volumen comercial (m3) de ??rboles en ??rea de Protecci??n dentro de la servidumbre del Predio  '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro8',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro8');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Cient??fico',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('0-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if($report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            }elseif(!$report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            }elseif( $numero==3)
            {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
          }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText(ucfirst($item->family->name),['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText(ucfirst($item->common->first()->name),['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                if($report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if($report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                if($report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                if($report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                if($report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                 if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                if($report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                if($report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                if($report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero==3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns9+=$ochenta;
                }
                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total General',['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);


    }

    public function boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero = 0)
    {
        /**
         * Cuadro 7
         */
        $section->addTextBreak();// Salto de l??nea
        if ($report && $numero == 0) {
            $section->addText('Cuadro 7. Distribuci??n diam??trica (cm) del volumen total (m3) de ??rboles ubicados en ??rea de Protecci??n del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 7. Distribuci??n diam??trica (cm) del volumen total (m3) de ??rboles en ??rea de Protecci??n dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }elseif ( $numero == 3) {
            $section->addText('Cuadro 14. Distribuci??n diam??trica (cm) del volumen total (m3) de ??rboles en ??rea de Protecci??n fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro7',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro7');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('0-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if($report && $numero == 0){
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }elseif(!$report && $numero == 0){
                $total= round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }elseif( $numero == 3){
                $total= round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if($report && $numero == 0){
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $diez= round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $diez= round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if($report && $numero == 0){
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $veinte= round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $veinte= round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                if($report && $numero == 0){
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $treinta= round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $treinta= round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                if($report && $numero == 0){
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $cuarenta= round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $cuarenta= round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                if($report && $numero == 0){
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $cincuenta= round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $cincuenta= round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                if($report && $numero == 0){
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $sesenta= round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $sesenta= round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                if($report && $numero == 0){
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $setenta= round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $setenta= round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                if($report && $numero == 0){
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $ochenta= round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $ochenta= round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column9+=$ochenta;
                }
                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal Comercial',['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);


        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $box2=ScientificName::where('commercial',0)->orderBy('name','ASC')->get();
        foreach ($box2 AS $item) {
            if($report && $numero == 0){
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }elseif(!$report && $numero == 0){
                $total= round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }elseif( $numero == 3){
                $total= round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if($report && $numero == 0){
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $diez= round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $diez= round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if($report && $numero == 0){
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $veinte= round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $veinte= round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                if($report && $numero == 0){
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $treinta= round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $treinta= round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                if($report && $numero == 0){
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $cuarenta= round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $cuarenta= round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                if($report && $numero == 0){
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $cincuenta= round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $cincuenta= round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                if($report && $numero == 0){
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $sesenta= round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $sesenta= round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                if($report && $numero == 0){
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $setenta= round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $setenta= round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                if($report && $numero == 0){
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif(!$report && $numero == 0){
                    $ochenta= round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }elseif( $numero == 3){
                    $ochenta= round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns9+=$ochenta;
                }
                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal No Comercial',['vMerge'          =>true,'bold'=>true,'vAlign'=>'both',
                                                                'size'            =>8,'unit'=>TblWidth::TWIP,
                                                                'borderLeftColor' =>'FFFFFF',
                                                                'borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total General',['vMerge'         =>true,'bold'=>true,'vAlign'=>'both',
                                                        'size'           =>8,'unit'=>TblWidth::TWIP,
                                                        'borderLeftColor'=>'FFFFFF','borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2 + $columns2,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column3 + $columns3,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column4 + $columns4,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column5 + $columns5,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column6 + $columns6,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column7 + $columns7,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column8 + $columns8,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column9 + $columns9,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column12 + $columns12,['bold'       =>true,'size'=>8,
                                                               'exactHeight'=>600]);

    }

    public function boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero=0)
    {
        /**
         * Cuadro 6
         */
        $section->addTextBreak();// Salto de l??nea
        if ($report && $numero==0) {
            $section->addText('Cuadro 6. Distribuci??n diam??trica (cm) del n??mero de ??rboles ubicados en ??rea de Protecci??n del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero==0) {
            $section->addText('Cuadro 6. Distribuci??n diam??trica (cm) del n??mero de ??rboles en ??rea de Protecci??n dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }elseif ( $numero==3) {
            $section->addText('Cuadro 13. Distribuci??n diam??trica (cm) del n??mero de ??rboles en ??rea de Protecci??n fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro 6',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 6');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('10-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $column1=0;
        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column10=0;
        $column11=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if ($report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }elseif ($numero==3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if ($report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if ($report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                if ($report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                if ($report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                if ($report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                if ($report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                if ($report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                if ($report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal Comercial',['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);


        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $box2=ScientificName::where('commercial',0)->orderBy('name','ASC')->get();
        foreach ($box2 AS $item) {
            if ($report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }elseif ($numero==3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if ($report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if ($report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                if ($report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                if ($report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                if ($report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                if ($report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                if ($report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                if ($report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal No Comercial',['vMerge'          =>true,'bold'=>true,'vAlign'=>'both',
                                                                'size'            =>8,'unit'=>TblWidth::TWIP,
                                                                'borderLeftColor' =>'FFFFFF',
                                                                'borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total General',['vMerge'         =>true,'bold'=>true,'vAlign'=>'both',
                                                        'size'           =>8,'unit'=>TblWidth::TWIP,
                                                        'borderLeftColor'=>'FFFFFF','borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2 + $columns2,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column3 + $columns3,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column4 + $columns4,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column5 + $columns5,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column6 + $columns6,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column7 + $columns7,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column8 + $columns8,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column9 + $columns9,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column12 + $columns12,['bold'       =>true,'size'=>8,
                                                               'exactHeight'=>600]);

    }

    public function boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero=0)
    {

        /**
         * Cuadro 5
         */

        if ($report && $numero==0) {
            $section->addText('Cuadro 5. Distribuci??n diam??trica (cm) del volumen comercial (m3) de los ??rboles en el Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero==0) {
            $section->addText('Cuadro 5. Distribuci??n diam??trica (cm) del volumen comercial (m3) de los ??rboles dentro de la servidumbre del Predio  '.$sysconf->id_predio);
        }elseif ($numero==3) {
            $section->addText('Cuadro 12. Distribuci??n diam??trica (cm) del volumen comercial (m3) de los ??rboles fuera de la servidumbre del Predio  '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro 5',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 5');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('0-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $column1=0;
        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if ($report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero==0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            }elseif ($numero==3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if ($report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[0, 19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if ($report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20, 29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }

                if ($report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                if ($report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                if ($report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                if ($report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                if ($report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                if ($report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero==0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }elseif ($numero==3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column9+=$ochenta;
                }
                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total General',['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);
    }

    public function boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero=0)
    {
        /**
         * Cuadro 4
         */
        if ($report && $numero == 0) {
            $section->addText('Cuadro 4. Distribuci??n diam??trica (cm) del volumen total (m3) de los ??rboles del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 4. Distribuci??n diam??trica (cm) del volumen total (m3) de los ??rboles dentro del servidumbre en el Predio  '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 11. Distribuci??n diam??trica (cm) del volumen total (m3) de los ??rboles fuera de la servidumbre en el Predio  '.$sysconf->id_predio);
        }


        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro4',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro4');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('0-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if ($report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                if ($report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal Comercial',['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);


        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $box2=ScientificName::where('commercial',0)->orderBy('name','ASC')->get();
        foreach ($box2 AS $item) {
            if ($report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de l??nea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                if ($report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal No Comercial',['vMerge'          =>true,'bold'=>true,'vAlign'=>'both',
                                                                'size'            =>8,'unit'=>TblWidth::TWIP,
                                                                'borderLeftColor' =>'FFFFFF',
                                                                'borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total General',['vMerge'         =>true,'bold'=>true,'vAlign'=>'both',
                                                        'size'           =>8,'unit'=>TblWidth::TWIP,
                                                        'borderLeftColor'=>'FFFFFF','borderRightColor'=>'FFFFFF']);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2 + $columns2,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column3 + $columns3,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column4 + $columns4,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column5 + $columns5,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column6 + $columns6,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column7 + $columns7,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column8 + $columns8,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column9 + $columns9,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column12 + $columns12,['bold'       =>true,'size'=>8,
                                                               'exactHeight'=>600]);

    }

    public function boxWordOne($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$styleFirstColumn)
    {
        $section->addText('Cuadro 1. Resumen de los individuos encontrados en el predio '.$sysconf->id_predio);
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro 1',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 1');
        $table->addRow(5);// Altura de l??nea 400
        $table->addCell(2000)->addText('Valor Comercial',['bold'       =>true,'size'=>12,
                                                          'exactHeight'=>300]);
        $table->addCell(2000)->addText('Ubicaci??n',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Individuos',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Sum g (m2)',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Vt (m3)',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Vc (m3)',['bold'=>true,'size'=>12]);

        $fuera=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Fuera')->count(),0);
        $dentro=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Dentro')->count(),0);

        $table->addRow(50);// Altura de l??nea 400
        $table->addCell(1000)->addText('Comercial',['vAlign'=>'center','vMerge'=>true,'size'=>12]);
        $table->addCell(3000)->addText('Fuera de AP',['vAlign'=>'center','size'=>12]);
        $table->addCell(500)->addText($fuera,['vAlign'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Fuera')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                                                              'size'     =>12,
                                                                                                                                                                              'width'    =>600,
                                                                                                                                                                              'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Fuera')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Fuera')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);

        $table->addRow(50,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(1000)->addText('Comercial',['alignment'=>'center','size'=>12]);
        $table->addCell(3000)->addText('Dentro de AP',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText($dentro,['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Dentro')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Dentro')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                                                                'size'     =>12,
                                                                                                                                                                                'width'    =>600,
                                                                                                                                                                                'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Dentro')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                                                                'size'     =>12,
                                                                                                                                                                                'width'    =>600,
                                                                                                                                                                                'color'    =>'000000']);


        $table->addRow(400,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(4000)->addText('Subtotal CO',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText('',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->count(),0),['alignment'=>'center',
                                                                                                                                         'size'     =>12,
                                                                                                                                         'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                            'size'     =>12,
                                                                                                                                            'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                             'size'     =>12,
                                                                                                                                             'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                             'size'     =>12,
                                                                                                                                             'width'    =>600]);

        $fuera=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Fuera')->count(),0);
        $dentro=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Dentro')->count(),0);


        $table->addRow(400,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(4000)->addText('No Comercial',['alignment'=>'center','size'=>12]);
        $table->addCell(3000)->addText('Fuera de AP',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText($fuera,['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Fuera')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                                                              'size'     =>12,
                                                                                                                                                                              'width'    =>600,
                                                                                                                                                                              'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Fuera')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Fuera')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);


        $table->addRow(50,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(4000)->addText('No Comercial',['alignment'=>'center','size'=>12]);
        $table->addCell(2000)->addText('Dentro de AP',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText($dentro,['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Dentro')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                                                               'size'     =>12,
                                                                                                                                                                               'width'    =>600,
                                                                                                                                                                               'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Dentro')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                                                                'size'     =>12,
                                                                                                                                                                                'width'    =>600,
                                                                                                                                                                                'color'    =>'000000']);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->where('protection_area','Dentro')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                                                                'size'     =>12,
                                                                                                                                                                                'width'    =>600,
                                                                                                                                                                                'color'    =>'000000']);


        $table->addRow(400,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal NC',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText('',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->count(),0),['alignment'=>'center',
                                                                                                                                         'size'     =>12,
                                                                                                                                         'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->sum('g_m'),2),['alignment'=>'center',
                                                                                                                                            'size'     =>12,
                                                                                                                                            'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->sum('vt_m'),2),['alignment'=>'center',
                                                                                                                                             'size'     =>12,
                                                                                                                                             'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','NC')->sum('vc_m'),2),['alignment'=>'center',
                                                                                                                                             'size'     =>12,
                                                                                                                                             'width'    =>600]);


        $table->addRow(400,$styleFirstColumn);// Altura de l??nea 400
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>12]);
        $table->addCell(500)->addText('',['alignment'=>'center','size'=>12]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->count(),0),['bold'     =>true,
                                                                                                               'alignment'=>'center',
                                                                                                               'size'     =>12,
                                                                                                               'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->sum('g_m'),2),['bold'     =>true,
                                                                                                                  'alignment'=>'center',
                                                                                                                  'size'     =>12,
                                                                                                                  'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->sum('vt_m'),2),['bold'     =>true,
                                                                                                                   'alignment'=>'center',
                                                                                                                   'size'     =>12,
                                                                                                                   'width'    =>600]);
        $table->addCell(500)->addText(number_format(ForestDataBase::where('farm_id',$sysconf->id)->sum('vc_m'),2),['bold'     =>true,
                                                                                                                   'alignment'=>'center',
                                                                                                                   'size'     =>12,
                                                                                                                   'width'    =>600]);

    }

    public function boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero=0)
    {
        /**
         * Cuadro 2
         */
        if ($report && $numero == 0) {
            $section->addText('Cuadro 2. Distribuci??n diam??trica (cm) del n??mero de ??rboles en el Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('CCuadro 2 Distribuci??n diam??trica (cm) del n??mero de ??rboles dentro de la servidumbre del Predio '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 9. Distribuci??n diam??trica (cm) del n??mero de ??rboles fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro 2',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 2');
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(1000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(1000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>300]);
        $table->addCell(1000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('10-19',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(50)->addText('Total general',['bold'=>true,'size'=>9]);


        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if ($report && $numero == 0) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
            } elseif (!$report && $numero == 0) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
            } elseif ($numero == 3) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de l??nea 400
                $table->addCell(1000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>300]);
                $table->addCell(1000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>300]);
                $table->addCell(1000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>300]);
                if ($report && $numero == 0) {
                $diez=ForestDataBase::whereBetween('dap',[0, 19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $diez=ForestDataBase::whereBetween('dap',[0, 19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $diez=ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                    if ($diez == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($veinte == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($veinte,['alignment'=>'center','size'=>8,
                                                          'width'    =>300]);
                    $column3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($treinta == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($treinta,['alignment'=>'center','size'=>8,
                                                           'width'    =>300]);
                    $column4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>300]);
                    $column5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>300]);
                    $column6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $column7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $column8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $column9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(2000)->addText('Subtotal Comercial',['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);

        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $boxNotComercial2=ScientificName::where('commercial',0)->orderBy('name','ASC')->get();


        foreach ($boxNotComercial2 AS $item) {

            if ($report && $numero == 0) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
            } elseif (!$report && $numero == 0) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
            } elseif ($numero == 3) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de l??nea 400
                $table->addCell(1000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>300]);
                $table->addCell(1000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>300]);

                $table->addCell(1000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>300]);
                if ($report && $numero == 0) {
                    $diez=ForestDataBase::whereBetween('dap',[0, 19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $diez=ForestDataBase::whereBetween('dap',[0, 19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $diez=ForestDataBase::whereBetween('dap',[0,19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($diez == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $veinte=ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($veinte == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($veinte,['alignment'=>'center','size'=>8,
                                                          'width'    =>300]);
                    $columns3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $treinta=ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($treinta == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($treinta,['alignment'=>'center','size'=>8,
                                                           'width'    =>300]);
                    $columns4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $cuarenta=ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>300]);
                    $columns5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $cincuenta=ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>300]);
                    $columns6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $sesenta=ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $setenta=ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                } elseif (!$report && $numero == 0) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap');
                } elseif ($numero == 3) {
                    $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap');
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Subtotal No Comercial',['vMerge'          =>true,'bold'=>true,'vAlign'=>'both',
                                                                'size'            =>8,'unit'=>TblWidth::TWIP,
                                                                'borderLeftColor' =>'FFFFFF',
                                                                'borderRightColor'=>'FFFFFF']);
        $table->addCell(4000)->addText('');
        $table->addCell(4000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Total General',['vMerge'         =>true,'bold'=>true,'vAlign'=>'both',
                                                        'size'           =>8,'unit'=>TblWidth::TWIP,
                                                        'borderLeftColor'=>'FFFFFF','borderRightColor'=>'FFFFFF']);
        $table->addCell(4000)->addText('');
        $table->addCell(4000)->addText('');
        $table->addCell(2000)->addText($column2 + $columns2,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column3 + $columns3,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column4 + $columns4,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column5 + $columns5,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column6 + $columns6,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column7 + $columns7,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column8 + $columns8,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column9 + $columns9,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>300]);
        $table->addCell(2000)->addText($column12 + $columns12,['bold'       =>true,'size'=>8,
                                                               'exactHeight'=>300]);

        /**
         * fin cuadro 2
         */
    }


    public function boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero=0)
    {

        /**
         * Cuadro 3
         */
        if ($report && $numero == 0) {
            $section->addText('Cuadro 3 Distribuci??n diam??trica (cm) de la sumatoria de ??reas basim??tricas (m2) de los ??rboles del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 3 Distribuci??n diam??trica (cm) del ??rea basal (m2) de los ??rboles dentro de la servidumbre del Predio '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 10. Distribuci??n diam??trica (cm) del ??rea basal (m2) de los ??rboles fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de l??nea
        $section->addText('Distribuci??n diam??trica (cm)');
        $section->addTextBreak();// Salto de l??nea

        $phpWord->addTableStyle('Cuadro 3',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 3');
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(4000)->addText('Nombre Cient??fico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(4000)->addText('Nombre com??n',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('0-19',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('20-29',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('30-39',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('40-49',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('50-59',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('60-69',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('70-79',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('80>=',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Total general',['bold'=>true,'size'=>9]);


        $column2=0;
        $column3=0;
        $column4=0;
        $column5=0;
        $column6=0;
        $column7=0;
        $column8=0;
        $column9=0;
        $column12=0;

        $box2=ScientificName::where('commercial',1)->orderBy('name','ASC')->get();

        foreach ($box2 AS $item) {
            if ($report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de l??nea 400
                $table->addCell(4000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(4000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(4000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if ($report && $numero == 0) {
                $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10, 19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20, 29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30, 39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40, 49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50, 59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60, 69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $column12+=$total;

            }
        }
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Subtotal Comercial',['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(4000)->addText('');
        $table->addCell(4000)->addText('');
        $table->addCell(2000)->addText($column2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($column12,['bold'=>true,'size'=>8]);


        $columns2=0;
        $columns3=0;
        $columns4=0;
        $columns5=0;
        $columns6=0;
        $columns7=0;
        $columns8=0;
        $columns9=0;
        $columns12=0;

        $box2=ScientificName::where('commercial',0)->orderBy('name','ASC')->get();
        foreach ($box2 AS $item) {
            if ($report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif (!$report && $numero == 0) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de l??nea 400
                $table->addCell(4000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(4000)->addText(ucfirst(strtolower($item->name)),['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(4000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                if ($report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10,19])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $diez=round(ForestDataBase::whereBetween('dap',[10, 19])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                if ($report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20,29])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $veinte=round(ForestDataBase::whereBetween('dap',[20, 29])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                if ($report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30,39])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $treinta=round(ForestDataBase::whereBetween('dap',[30, 39])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                if ($report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40,49])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cuarenta=round(ForestDataBase::whereBetween('dap',[40, 49])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                if ($report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50,59])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $cincuenta=round(ForestDataBase::whereBetween('dap',[50, 59])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                if ($report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60,69])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $sesenta=round(ForestDataBase::whereBetween('dap',[60, 69])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                if ($report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $setenta=round(ForestDataBase::whereBetween('dap',[70,79])->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                if ($report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif (!$report && $numero == 0) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                } elseif ($numero == 3) {
                    $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                }
                if ($ochenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($ochenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns9+=$ochenta;
                }

                $table->addCell(2000)->addText($total,['bold'=>true,'size'=>8]);
                $columns12+=$total;

            }
        }
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Subtotal No Comercial',['vMerge'          =>true,'bold'=>true,'vAlign'=>'both',
                                                                'size'            =>8,'unit'=>TblWidth::TWIP,
                                                                'borderLeftColor' =>'FFFFFF',
                                                                'borderRightColor'=>'FFFFFF']);
        $table->addCell(4000)->addText('');
        $table->addCell(4000)->addText('');
        $table->addCell(2000)->addText($columns2,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns3,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns4,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns5,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns6,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns7,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns8,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns9,['bold'=>true,'size'=>8]);
        $table->addCell(2000)->addText($columns12,['bold'=>true,'size'=>8]);
        $table->addRow(400);// Altura de l??nea 400
        $table->addCell(4000)->addText('Total General',['vMerge'         =>true,'bold'=>true,'vAlign'=>'both',
                                                        'size'           =>8,'unit'=>TblWidth::TWIP,
                                                        'borderLeftColor'=>'FFFFFF','borderRightColor'=>'FFFFFF']);
        $table->addCell(4000)->addText('');
        $table->addCell(4000)->addText('');
        $table->addCell(2000)->addText($column2 + $columns2,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column3 + $columns3,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column4 + $columns4,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column5 + $columns5,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column6 + $columns6,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column7 + $columns7,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column8 + $columns8,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column9 + $columns9,['bold'       =>true,'size'=>8,
                                                             'exactHeight'=>600]);
        $table->addCell(2000)->addText($column12 + $columns12,['bold'       =>true,'size'=>8,
                                                               'exactHeight'=>600]);
        /**
         * fin cuadro 3
         */
    }

    public function importFamily(Request $request)
    {
        $file=$request->file();

        $import=new ImportFamilyExcel;

        Excel::import($import,$file['data_base']);
        return redirect()->back();
    }
}
