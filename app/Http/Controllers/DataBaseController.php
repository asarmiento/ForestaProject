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
use Illuminate\Support\Facades\Log;
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
        if (ForestDataBase::where('farm_id',$request->get('farm_id'))->where('year',Carbon::now()->year)->count() > 0 && $request->get('reeplay') == 'on') {
            DB::statement("DELETE FROM `forest_database` WHERE farm_id = ".$request->get('farm_id')." AND year =".Carbon::now()->year);
            //ForestDataBase::where('farm_id',$this->farm)->where('year',Carbon::now()->year)->get()->destroy();
        }
        $file=$request->file();
        $import=new ImportNewExcelDataBase($request->get('farm_id'));

        Excel::import($import,$file['data_base']);
        return redirect()->back();
    }


    public function reportWord(Request $request)
    {
        $sysconf=Farm::find($request->get('farm_id'));
        Log::info(json_encode($sysconf));
        $phpWord=new PhpWord;
        // Establecer el estilo predeterminado
        $phpWord->setDefaultFontName('Times New Roma');                              //Fuente
        $phpWord->setDefaultFontSize(16);                                            //Tamaño de fuente
        $phpWord->setDefaultParagraphStyle(['vAlign'=>'center']);                    //Tamaño de fuente

        //Añadir página
        $section=$phpWord->createSection(['orientation'=>'landscape']);

        // Agregar directorio
        $styleTOC=['tabLeader'=>\PhpOffice\PhpWord\Style\TOC::TABLEADER_DOT];
        $styleFont=['spaceAfter'=>60,'name'=>'Times New Roma','size'=>10];
        $section->addTOC($styleFont,$styleTOC);
        /**
         *
         * //Estilo por Defecto
         * $section->addText('Proyecto SIEPAC',['size'=>16,'bold'=>true]);
         * $section->addTextBreak();// Salto de línea
         * $section->addText('Inventario de Predios del Tramo 16',['size'=>16]);
         * $section->addTextBreak();// Salto de línea
         *
         *
         * $section->addText('Informe del Predio '.$sysconf->id_predio,['size'=>14,'bold'=>true]);
         * $section->addTextBreak();// Salto de línea
         * $section->addText($sysconf->owner,['size'=>14,'bold'=>true]);
         * $section->addTextBreak();// Salto de línea
         * $section->addTextBreak();// Salto de línea
         * $section->addTextBreak();// Salto de línea
         * $section->addText('Elaborado por:',['size'=>14]);
         * $section->addText('Ing. Álvaro Sibaja Villegas',['size'=>14]);
         * $section->addTextBreak(7);// Salto de línea
         *
         *
         * $section->addText('San José, Costa Rica ',['size'=>14]);
         *
         * $section->addTextBreak(5);// Salto de línea
         *
         * $section->addText('1. Consideraciones generales',['bold'=>true]);
         * $section->addTextBreak();// Salto de línea
         * $fontStyle=new \PhpOffice\PhpWord\Style\Font;
         * $fontStyle->setAuto(true);
         * $fontStyle->setName('Tahoma');
         * $fontStyle->setSize(13);
         * $myTextElement=$section->addText('Los datos base para la elaboración del presente informe se obtuvieron
         * en primer lugar de la información proporcionada por el Ing. Allan Montoya, que comprende los shapefile de los predios, líneas de tensión, torres y servidumbres; en segundo lugar, se necesitó recolectar datos de campo para conocer la ubicación y características de las especies dentro y fuera de la servidumbre.  En el inventario se ubicaron todos los árboles con un diámetro normal mayor a 15 cm, los cuales se georreferenciaron con un GPS Garmin 64s.');
         * $myTextElement->setFontStyle($fontStyle);
         *
         *
         * $section->addText('A continuación, se presenta la nomenclatura utilizada en el presente informe.');
         *
         *
         * $nameScien = ForestDataBase::where('farm_id',$sysconf->id)->groupBy('name_cientifict')->count();
         * $section->addText('El inventario forestal registró    '.$nameScien.'
         * especies diferentes, distribuidas en    10
         * familias. Un total de
         * 71    individuos fueron medidos, de estos    58     pertenecientes a especies con valor comercial,
         * mientras que    13    a especies sin valor comercial. Las áreas basimétricas (g) sumaron en total
         * 6.754    m2,  además, se encontró un volumen comercial (Vc) de    6.414    m3 y el volumen total
         * (Vt) alcanzó    56.269    m3 (Cuadro 1).
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
        // Estilo de primera línea
        $styleFirstColumn=[
            'borderSize'       =>1,
            'borderTopColor'   =>'FFFFFF',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'vAlign'           =>'center'];
        // Estilo de primera línea

        /**
         * cuadro 1
         */
        $this->boxWordOne($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$styleFirstColumn);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 2
         */
        $this->boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 3
         */
        $this->boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 4
         */
        $this->boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 5
         */
        $this->boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 6
         */
        $this->boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 7
         */
        $this->boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 8
         */
        $this->boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 8
         */
        $this->boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord);
        $section->addTextBreak();// Salto de línea

        // El documento generado es Word2007
        $writer=\PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
        $writer->save('reporte_'.$sysconf->id_predio.'_uno.docx');
        return redirect()->back();

    }

    public function reportWordTwo(Request $request)
    {
        $sysconf=Farm::find($request->get('farm_id'));

        $phpWord=new PhpWord;
        // Establecer el estilo predeterminado
        $phpWord->setDefaultFontName('Times New Roma');                              //Fuente
        $phpWord->setDefaultFontSize(16);                                            //Tamaño de fuente
        $phpWord->setDefaultParagraphStyle(['vAlign'=>'center']);                    //Tamaño de fuente

        //Añadir página
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
        // Estilo de primera línea
        $styleFirstColumn=[
            'borderSize'       =>1,
            'borderTopColor'   =>'FFFFFF',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'vAlign'           =>'center'];
        // Estilo de primera línea

        /**
         * cuadro 1
         */
        $this->boxWordOne($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$styleFirstColumn);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 2
         */
        $this->boxWordTwo($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 3
         */
        $this->boxWordThree($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 4
         */
        $this->boxWordFour($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 5
         */
        $this->boxWordFive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 6
         */
        $this->boxWordSix($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 7
         */
        $this->boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 8
         */
        $this->boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea
        /**
         * cuadro 8
         */
        $this->boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,false,3);
        $section->addTextBreak();// Salto de línea

        // El documento generado es Word2007
        $writer=\PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
        $writer->save('reporte_'.$sysconf->id_predio.'_dos.docx');

        return redirect()->back();
    }

    public function boxWordNive($sysconf,$section,$styleTable,$styleFirstRow,$phpWord)
    {
        $section->addText('Anexo 1. Árboles inventariados en el Predio  '.$sysconf->id_predio);
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Anexo1',$styleTable,$styleFirstRow);


        $table=$section->addTable('Anexo1');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(500)->addText('Vano',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Árbol',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Nombre Científico',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Nombre común',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Cobertura',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Comercial',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Servidumbre',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Área de Protección',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Dap (cm)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Ht (m)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Hc (m)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('G (m²)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Vt (m²)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Vc (m²)',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Coord X',['bold'=>true,'size'=>9]);
        $table->addCell(500)->addText('Coord Y',['bold'=>true,'size'=>9]);


        $datas=ForestDataBase::where('farm_id',$sysconf->id)->orderBy('name_cientifict','ASC')->get();

        foreach ($datas AS $data) {
            $table->addRow(5);// Altura de línea 400
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

    public function boxWordEight($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true)
    {
        if ($report) {
            $section->addText('Cuadro 8. Distribución diamétrica (cm) del volumen comercial (m3) de árboles ubicados en Área de Protección del Predio '.$sysconf->id_predio);
        } else {
            $section->addText('Cuadro 8. Distribución diamétrica (cm) del volumen comercial (m3) de árboles en Áreas de Protección dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro8',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro8');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Científico',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
            $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
            if ($total > 0) {
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                $diez=round(ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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

    public function boxWordSeven($sysconf,$section,$styleTable,$styleFirstRow,$phpWord,$report=true,$numero = 0)
    {
        /**
         * Cuadro 7
         */
        $section->addTextBreak();// Salto de línea
        if ($report && $numero == 0) {
            $section->addText('Cuadro 7. Distribución diamétrica (cm) del volumen total (m3) de árboles ubicados en Área de Protección del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 7. Distribución diamétrica (cm) del volumen total (m3) de árboles en Área de Protección dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }elseif ( $numero == 3) {
            $section->addText('Cuadro 14. Distribución diamétrica (cm) del volumen total (m3) de árboles en Área de Protección fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro7',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro7');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
            }else{
                $total= round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
            $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }elseif($numero ==3){
            $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
        $table->addRow(5);// Altura de línea 400
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
        $section->addTextBreak();// Salto de línea
        if ($report && $numero==0) {
            $section->addText('Cuadro 6. Distribución diamétrica (cm) del número de árboles ubicados en Área de Protección del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero==0) {
            $section->addText('Cuadro 6. Distribución diamétrica (cm) del número de árboles en Área de Protección dentro de la servidumbre del Predio '.$sysconf->id_predio);
        }elseif ( $numero==3) {
            $section->addText('Cuadro 13. Distribución diamétrica (cm) del número de árboles en Área de Protección fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro 6',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 6');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }elseif ($numero==3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                $diez=round(ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }elseif ($numero==3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
            }if ($total > 0) {
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                $diez=round(ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->count('dap'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
        $table->addRow(5);// Altura de línea 400
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
            $section->addText('Cuadro 5. Distribución diamétrica (cm) del volumen comercial (m3) de los árboles en el Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero==0) {
            $section->addText('Cuadro 5. Distribución diamétrica (cm) del volumen comercial (m3) de los árboles dentro de la servidumbre del Predio  '.$sysconf->id_predio);
        }elseif ($numero==3) {
            $section->addText('Cuadro 12. Distribución diamétrica (cm) del volumen comercial (m3) de los árboles fuera de la servidumbre del Predio  '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro 5',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 5');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                $diez=round(ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vc_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
            $section->addText('Cuadro 4. Distribución diamétrica (cm) del volumen total (m3) de los árboles del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 4. Distribución diamétrica (cm) del volumen total (m3) de los árboles dentro del servidumbre en el Predio  '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 11. Distribución diamétrica (cm) del volumen total (m3) de los árboles fuera de la servidumbre en el Predio  '.$sysconf->id_predio);
        }


        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro4',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro4');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(2000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(2000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
                $table->addRow(5);// Altura de línea 400
                $table->addCell(2000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(2000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(2000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);
                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('vt_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(5);// Altura de línea 400
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
        $table->addRow(5);// Altura de línea 400
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
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro 1',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 1');
        $table->addRow(5);// Altura de línea 400
        $table->addCell(2000)->addText('Valor Comercial',['bold'       =>true,'size'=>12,
                                                          'exactHeight'=>300]);
        $table->addCell(2000)->addText('Ubicación',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Individuos',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Sum g (m2)',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Vt (m3)',['bold'=>true,'size'=>12]);
        $table->addCell(2000)->addText('Vc (m3)',['bold'=>true,'size'=>12]);

        $fuera=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Fuera')->count(),0);
        $dentro=number_format(ForestDataBase::where('farm_id',$sysconf->id)->where('commercial','CO')->where('protection_area','Dentro')->count(),0);

        $table->addRow(50);// Altura de línea 400
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

        $table->addRow(50,$styleFirstColumn);// Altura de línea 400
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


        $table->addRow(400,$styleFirstColumn);// Altura de línea 400
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


        $table->addRow(400,$styleFirstColumn);// Altura de línea 400
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


        $table->addRow(50,$styleFirstColumn);// Altura de línea 400
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


        $table->addRow(400,$styleFirstColumn);// Altura de línea 400
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


        $table->addRow(400,$styleFirstColumn);// Altura de línea 400
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
            $section->addText('Cuadro 2. Distribución diamétrica (cm) del número de árboles en el Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('CCuadro 2 Distribución diamétrica (cm) del número de árboles dentro de la servidumbre del Predio '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 9. Distribución diamétrica (cm) del número de árboles fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro 2',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 2');
        $table->addRow(400);// Altura de línea 400
        $table->addCell(1000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(1000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>300]);
        $table->addCell(1000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count();
            } elseif ($numero == 3) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count();
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de línea 400
                $table->addCell(1000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>300]);
                $table->addCell(1000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>300]);
                $table->addCell(1000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>300]);

                $diez=ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($diez == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($veinte == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($veinte,['alignment'=>'center','size'=>8,
                                                          'width'    =>300]);
                    $column3+=$veinte;
                }
                $treinta=ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($treinta == 0) {
                    $table->addCell(50)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(50)->addText($treinta,['alignment'=>'center','size'=>8,
                                                           'width'    =>300]);
                    $column4+=$treinta;
                }
                $cuarenta=ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>300]);
                    $column5+=$cuarenta;
                }
                $cincuenta=ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>300]);
                    $column6+=$cincuenta;
                }
                $sesenta=ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $column7+=$sesenta;
                }
                $setenta=ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $column8+=$setenta;
                }
                $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
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
        $table->addRow(400);// Altura de línea 400
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
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Dentro')->where('farm_id',$sysconf->id)->count();
            } elseif ($numero == 3) {
                $total=ForestDataBase::where('name_cientifict',$item->name)->where('servitude','Fuera')->where('farm_id',$sysconf->id)->count();
            }

            if ($total > 0) {

                $table->addRow(400);// Altura de línea 400
                $table->addCell(50)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                  'exactHeight'=>300]);
                $table->addCell(50)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                          'width' =>300]);
                $table->addCell(50)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                           'size'  =>9,
                                                                           'width' =>300]);

                $diez=ForestDataBase::whereBetween('dap',[0,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                $veinte=ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>300]);
                    $columns3+=$veinte;
                }
                $treinta=ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns4+=$treinta;
                }
                $cuarenta=ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>300]);
                    $columns5+=$cuarenta;
                }
                $cincuenta=ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>300]);
                    $columns6+=$cincuenta;
                }
                $sesenta=ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns7+=$sesenta;
                }
                $setenta=ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>300]);
                    $columns8+=$setenta;
                }
                $ochenta=ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->count();
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
        $table->addRow(400);// Altura de línea 400
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
        $table->addRow(400);// Altura de línea 400
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
            $section->addText('Cuadro 3 Distribución diamétrica (cm) de la sumatoria de áreas basimétricas (m2) de los árboles del Predio '.$sysconf->id_predio);
        } elseif (!$report && $numero == 0) {
            $section->addText('Cuadro 3 Distribución diamétrica (cm) del área basal (m2) de los árboles dentro de la servidumbre del Predio '.$sysconf->id_predio);
        } elseif ($numero == 3) {
            $section->addText('Cuadro 10. Distribución diamétrica (cm) del área basal (m2) de los árboles fuera de la servidumbre del Predio '.$sysconf->id_predio);
        }
        $section->addTextBreak();// Salto de línea
        $section->addText('Distribución diamétrica (cm)');
        $section->addTextBreak();// Salto de línea

        $phpWord->addTableStyle('Cuadro 3',$styleTable,$styleFirstRow);

        $table=$section->addTable('Cuadro 3');
        $table->addRow(400);// Altura de línea 400
        $table->addCell(4000)->addText('Families',['bold'=>true,'size'=>9]);
        $table->addCell(4000)->addText('Nombre Científico',['bold' =>true,'size'=>9,
                                                            'width'=>600]);
        $table->addCell(4000)->addText('Nombre común',['bold'=>true,'size'=>9]);
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
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de línea 400
                $table->addCell(4000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(4000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(4000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $column2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $column3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $column5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $column6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $column8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(400);// Altura de línea 400
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
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Dentro')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            } elseif ($numero == 3) {
                $total=round(ForestDataBase::where('name_cientifict',$item->name)->where('protection_area','Fuera')->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
            }
            if ($total > 0) {
                $table->addRow(400);// Altura de línea 400
                $table->addCell(4000)->addText($item->family->name,['vAlign'     =>'center','size'=>8,
                                                                    'exactHeight'=>600]);
                $table->addCell(4000)->addText($item->name,['vAlign'=>'center','size'=>8,
                                                            'width' =>600]);
                $table->addCell(4000)->addText($item->common->first()->name,['vAlign'=>'center',
                                                                             'size'  =>9,
                                                                             'width' =>600]);

                $diez=round(ForestDataBase::whereBetween('dap',[10,
                    19])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($diez == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($diez,['alignment'=>'center','size'=>8]);
                    $columns2+=$diez;
                }
                $veinte=round(ForestDataBase::whereBetween('dap',[20,
                    29])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($veinte == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($veinte,['alignment'=>'center','size'=>8,
                                                            'width'    =>600]);
                    $columns3+=$veinte;
                }
                $treinta=round(ForestDataBase::whereBetween('dap',[30,
                    39])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($treinta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($treinta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns4+=$treinta;
                }
                $cuarenta=round(ForestDataBase::whereBetween('dap',[40,
                    49])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($cuarenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cuarenta,['alignment'=>'center','size'=>8,
                                                              'width'    =>600]);
                    $columns5+=$cuarenta;
                }
                $cincuenta=round(ForestDataBase::whereBetween('dap',[50,
                    59])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($cincuenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($cincuenta,['alignment'=>'center','size'=>8,
                                                               'width'    =>600]);
                    $columns6+=$cincuenta;
                }
                $sesenta=round(ForestDataBase::whereBetween('dap',[60,
                    69])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($sesenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($sesenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns7+=$sesenta;
                }
                $setenta=round(ForestDataBase::whereBetween('dap',[70,
                    79])->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
                if ($setenta == 0) {
                    $table->addCell(2000)->addText('',['alignment'=>'center','size'=>8]);
                } else {
                    $table->addCell(2000)->addText($setenta,['alignment'=>'center','size'=>8,
                                                             'width'    =>600]);
                    $columns8+=$setenta;
                }
                $ochenta=round(ForestDataBase::where('dap','>=',80)->where('name_cientifict',$item->name)->where('farm_id',$sysconf->id)->sum('g_m'),3,PHP_ROUND_HALF_UP);
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
        $table->addRow(400);// Altura de línea 400
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
        $table->addRow(400);// Altura de línea 400
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
