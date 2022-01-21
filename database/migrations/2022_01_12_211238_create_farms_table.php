<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateFarmsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('farms', function (Blueprint $table) {
            $table->id();
            $table->string('order')->nullable();
            $table->string('code')->nullable();
            $table->string('office_sinac')->nullable();
            $table->string('count_vano')->nullable();
            $table->string('detail_vano')->nullable();
            $table->string('id_predio')->nullable();
            $table->string('owner')->nullable();
            $table->string('card')->nullable();
            $table->string('folio_real')->nullable();
            $table->string('plane')->nullable();
            $table->string('appointment_contract')->nullable();
            $table->timestamps();
            $table->engine='InnoDB';
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('farms');
    }
}
