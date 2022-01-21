<?php


namespace App\Models;


class ScientificName extends Entity
{
    protected $table = 'scientific_names';
    protected $fillable = ['name', 'commercial', 'family_id'];

    public function family()
    {
        return $this->belongsTo(Family::class,'family_id','id');
    }

    public function common()
    {
        return $this->hasMany(CommonName::class,'scientific_name_id','id');
    }
}
