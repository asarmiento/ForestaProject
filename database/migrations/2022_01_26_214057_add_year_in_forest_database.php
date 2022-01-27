<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddYearInForestDatabase extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('forest_database', function (Blueprint $table) {
           $table->year('year')->nullable()->after('farm_id');
        });
        Schema::table('farms', function (Blueprint $table) {
           $table->string('logitud_km')->nullable()->after('appointment_contract');
           $table->string('predio_before')->nullable()->after('logitud_km');
           $table->string('predio_after')->nullable()->after('predio_before');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('forest_database', function (Blueprint $table) {
            //
        });
    }
}
