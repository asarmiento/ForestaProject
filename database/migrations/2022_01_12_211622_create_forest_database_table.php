<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateForestDatabaseTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('forest_database', function (Blueprint $table) {
            $table->id();
            $table->string('farm_id');
            $table->string('vano');
            $table->string('tree');
            $table->string('family');
            $table->string('name_cientifict');
            $table->string('name_common');
            $table->string('coverage');
            $table->string('commercial');
            $table->string('servitude');
            $table->string('protection_area');
            $table->decimal('dap',20,2);
            $table->decimal('ht_m',20,2);
            $table->decimal('hc_m',20,2);
            $table->decimal('g_m',20,7);
            $table->decimal('vt_m',20,7);
            $table->decimal('vc_m',20,7);
            $table->decimal('coord_x',20,5);
            $table->decimal('coord_y',20,5);
            $table->timestamps();
            $table->engine = 'InnoDB';
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('forest_database');
    }
}
