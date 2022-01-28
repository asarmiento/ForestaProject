<?php


namespace App\Models;


class Family extends Entity
{
    protected $table ='families';
    protected $fillable =['name'];

    public function setNameAttribute($value)
    {
        $this->attributes['name'] = ucwords(strtolower($value));
    }
}
