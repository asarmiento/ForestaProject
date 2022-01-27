<?php


namespace App\Models;


class ForestDataBase extends Entity
{
    protected $table='forest_database';
    protected $fillable=['farm_id','year','vano','tree','family','name_cientifict','name_common','coverage','commercial',
        'servitude','protection_area','dap','ht_m','hc_m','g_m','vt_m','vc_m','coord_x','coord_y'];


}
