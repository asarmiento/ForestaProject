<?php

namespace App\Imports;

use App\Models\Farm;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToCollection;
use Maatwebsite\Excel\Concerns\WithHeadingRow;

class ImportNewExcelFarm implements ToCollection, WithHeadingRow
{
    /**
    * @param Collection $collection
    */
    public function collection(Collection $collection)
    {
        foreach ($collection AS $coll){
//echo json_encode($coll); die;
        $farm = new Farm();
        $farm->order =$coll['orden'];
        $farm->code  =$coll['id'];
        $farm->office_sinac =$coll['oficina_sinac'];
        $farm->count_vano  =$coll['cantidad_vanos'];
        $farm->detail_vano  =$coll['detalle_vanos'];
        $farm->id_predio  =$coll['id_predio'];
        $farm->owner  =$coll['propietario'];
        $farm->card  =$coll['cedula'];
        $farm->folio_real  =$coll['folio_real'];
        $farm->plane  =$coll['plano'];
        $farm->appointment_contract =$coll['cita_contrato_servidumbre'];
        $farm->save();
        }
    }
}
