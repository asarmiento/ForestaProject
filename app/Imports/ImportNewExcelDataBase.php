<?php

namespace App\Imports;

use App\Models\ForestDataBase;
use Carbon\Carbon;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Concerns\ToCollection;
use Maatwebsite\Excel\Concerns\WithHeadingRow;

class ImportNewExcelDataBase implements ToCollection,WithHeadingRow
{
    private $farm;

    public function __construct($farm)
    {

        $this->farm=$farm;

    }
    /**
     * @param Collection $collection
     */
    public function collection(Collection $collection)
    {


        foreach ($collection AS $coll) {
            //echo json_encode($coll); die;

          $farm=new ForestDataBase;

            $farm->farm_id=$this->farm;
            $farm->vano=$coll['vano'];
            $farm->year=Carbon::now()->year;
            $farm->tree=$coll['arbol'];
            $farm->family=$coll['familia'];
            $farm->name_cientifict=$coll['nombre_cientifico'];
            $farm->name_common=$coll['nombre_comun'];
            $farm->coverage=$coll['cobertura'];
            $farm->commercial=$coll['comercial'];
            $farm->servitude=$coll['servidumbre'];
            $farm->protection_area=$coll['area_proteccion'];
            $farm->dap=$coll['dap'];
            $farm->ht_m=$coll['ht'];
            $farm->hc_m=$coll['hc'];
            $farm->g_m=$coll['g'];
            $farm->vt_m=$coll['vt'];
            $farm->vc_m=$coll['vc'];
            $farm->coord_x=$coll['coord_x'];
            $farm->coord_y=$coll['coord_y'];
            $farm->save();
        }
    }
}
