<?php

namespace App\Imports;

use App\Models\CommonName;
use App\Models\Family;
use App\Models\ScientificName;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToCollection;
use Maatwebsite\Excel\Concerns\WithHeadingRow;

class ImportFamilyExcel implements ToCollection, WithHeadingRow
{
    /**
    * @param Collection $collection
    */
    public function collection(Collection $collection)
    {
        foreach ($collection AS $item){

            if(Family::where('name',$item['familia'])->count() == 0){
                $families = new Family();
                $families->name = $item['familia'];
                $families->save();
            }

          $family=  Family::where('name',$item['familia'])->first();
            if(ScientificName::where('name',$item['nombre_cientifico'])->where('family_id',$family->id)->count() == 0){
                $families = new ScientificName();
                $families->name = $item['nombre_cientifico'];
                $families->family_id=$family->id;
                if($item['comercial'] == 1){
                    $families->commercial=true;
                }else{
                    $families->commercial=false;
                }
                $families->save();
            }

          $scientifict = ScientificName::where('name',$item['nombre_cientifico'])->first();
            if(CommonName::where('name',$item['nombre_comun'])->where('scientific_name_id',$scientifict->id)->count() == 0){
                $families = new CommonName();
                $families->name = $item['nombre_comun'];
                $families->scientific_name_id=$scientifict->id;
                $families->save();
            }
        }
    }
}
