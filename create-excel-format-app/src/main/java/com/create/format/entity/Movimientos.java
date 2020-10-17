package com.create.format.entity;

public class Movimientos {
	
	private String fecha;
    private String fondo;
    private String descripcion;
    private String detalle;
    private Double numeroDeCuotas;
    private Double valorCuota;
    private Double saldo;
    private String codigoMoneda;
    private String tipoFondo;
    private String signo;
    
	public Movimientos(String fecha, String fondo, String descripcion, String detalle, Double numeroDeCuotas,
			Double valorCuota, Double saldo, String codigoMoneda, String tipoFondo, String signo) {
		super();
		this.fecha = fecha;
		this.fondo = fondo;
		this.descripcion = descripcion;
		this.detalle = detalle;
		this.numeroDeCuotas = numeroDeCuotas;
		this.valorCuota = valorCuota;
		this.saldo = saldo;
		this.codigoMoneda = codigoMoneda;
		this.tipoFondo = tipoFondo;
		this.signo = signo;
	}
	
	public String getFecha() {
		return fecha;
	}
	public void setFecha(String fecha) {
		this.fecha = fecha;
	}
	public String getFondo() {
		return fondo;
	}
	public void setFondo(String fondo) {
		this.fondo = fondo;
	}
	public String getDescripcion() {
		return descripcion;
	}
	public void setDescripcion(String descripcion) {
		this.descripcion = descripcion;
	}
	public String getDetalle() {
		return detalle;
	}
	public void setDetalle(String detalle) {
		this.detalle = detalle;
	}
	public Double getNumeroDeCuotas() {
		return numeroDeCuotas;
	}
	public void setNumeroDeCuotas(Double numeroDeCuotas) {
		this.numeroDeCuotas = numeroDeCuotas;
	}
	public Double getValorCuota() {
		return valorCuota;
	}
	public void setValorCuota(Double valorCuota) {
		this.valorCuota = valorCuota;
	}
	public Double getSaldo() {
		return saldo;
	}
	public void setSaldo(Double saldo) {
		this.saldo = saldo;
	}
	public String getCodigoMoneda() {
		return codigoMoneda;
	}
	public void setCodigoMoneda(String codigoMoneda) {
		this.codigoMoneda = codigoMoneda;
	}
	public String getTipoFondo() {
		return tipoFondo;
	}
	public void setTipoFondo(String tipoFondo) {
		this.tipoFondo = tipoFondo;
	}
	public String getSigno() {
		return signo;
	}
	public void setSigno(String signo) {
		this.signo = signo;
	}

}
