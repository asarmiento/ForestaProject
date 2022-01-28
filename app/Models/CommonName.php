<?php


namespace App\Models;


class CommonName extends Entity
{
    protected $table='common_names';
    protected $fillable=['name','scientific_name_id'];



    public function setNameAttribute($value)
    {
        $this->attributes['name'] = ucwords(strtolower($value));
    }

    public function scientific()
    {
        return $this->belongsTo(ScientificName::class,'scientific_name_id','id');
    }
}
