<?php


namespace App\Models;


class CommonName extends Entity
{
protected $table = 'common_names';
protected $fillable =['name', 'scientific_name_id'];
}
