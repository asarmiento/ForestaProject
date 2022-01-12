@extends('layouts.app')

@section('content')
<div id="app" class="container">
    <example-component></example-component>
</div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
    @endsection
