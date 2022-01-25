<?php


namespace App\Models;


class Farm extends Entity
{
    protected $table = 'farms';
    protected $fillable =[ 'order', 'code', 'office_sinac', 'count_vano', 'detail_vano', 'id_predio', 'owner', 'card', 'folio_real', 'plane', 'appointment_contract'];

    public function database()
    {
        return $this->hasMany(ForestDataBase::class,'farm_id','id');
    }
}
