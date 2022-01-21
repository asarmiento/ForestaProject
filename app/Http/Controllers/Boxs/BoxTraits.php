<?php


namespace App\Http\Controllers\Boxs;


use App\Models\ForestDataBase;
use App\Models\ScientificName;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\TblWidth;

trait BoxTraits
{


    public function boxTwo($request)
    {
        $phpWord=new PhpWord;
        // Establecer el estilo predeterminado
        $phpWord->setDefaultFontName('Canción de imitación');//Fuente
        $phpWord->setDefaultFontSize(16);                    //Tamaño de fuente

        //Añadir página
        $section=$phpWord->createSection();

        // Agregar directorio
        $styleTOC=['tabLeader'=>\PhpOffice\PhpWord\Style\TOC::TABLEADER_DOT];
        $styleFont=['spaceAfter'=>60,'name'=>'Times','size'=>10];
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
            'width'            =>600,
            'valign'           =>'center',
        ];
        $styleFirstRow=[
            'borderSize'       =>1,
            'borderTopColor'   =>'1A446C',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'valign'           =>'center'];
        // Estilo de primera línea
        $styleFirstColumn=[
            'borderSize'       =>1,
            'borderTopColor'   =>'FFFFFF',
            'borderRightColor' =>'FFFFFF',
            'borderLeftColor'  =>'FFFFFF',
            'borderBottomColor'=>'1A446C',
            'textDirection'    =>'tbRl',
            'valign'           =>'center'];

return $section;
    }
}
