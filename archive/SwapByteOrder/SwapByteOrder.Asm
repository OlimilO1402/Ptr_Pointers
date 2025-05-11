;#########################################################################
;#		Author: 	Kai Liebenau
;#		Email: 		Kai@ActiveVB.de
;#		Publiziert: http://www.activevb.de/rubriken/komponenten/dll31/dll31.html
;#		Erstellt:	07.08.2008
;#
;#		Assembler directives
;#########################################################################
.686
.model flat,stdcall
option casemap:none

SAFEARRAYBOUND struc dword
  cElements   dword     ?
  lLbound     sdword    ?
SAFEARRAYBOUND  ends
LPSAFEARRAYBOUND  typedef ptr SAFEARRAYBOUND

SAFEARRAY   struc dword
  cDims       word      ?
  fFeatures   word      ?
  cbElements  dword     ?
  cLocks      dword     ?
  pvData      dword     ?
  rgsabound   SAFEARRAYBOUND  1t DUP (<>)
SAFEARRAY ends
LPSAFEARRAY typedef ptr SAFEARRAY

ERROR_INVALID_PARAMETER equ 87
FADF_HAVEVARTYPE	equ 80h

include windows.inc

.code

;#########################################################################
;# 		Einstiegsprozedur für die DLL
;#########################################################################
Align sizeof(DWord)
DllEntry proc hInst:HINSTANCE, reason:DWORD, reserved1:DWORD
	mov eax, TRUE
	ret
DllEntry Endp

;#########################################################################
;# 		Input: 	Zeiger auf eine 16Bit Variable
;# 		Output:	nichts
;#########################################################################
OPTION EPILOGUE:NONE
OPTION PROLOGUE:NONE
Align sizeof(DWord)
SwapByteOrder16	proc	lpValue:ptr WORD
	
	mov eax, [esp+4];lpValue
	.If eax
		mov cx, Word ptr [eax]
		xchg ch, cl
		mov Word ptr [eax], cx 
	.endif	
	ret 4
SwapByteOrder16 endp

;#########################################################################
;# 		Input: 	Zeiger auf eine 32Bit Variable
;# 		Output:	nichts
;#########################################################################
Align sizeof(DWord)
SwapByteOrder32	proc	lpValue:ptr DWORD
	
	mov eax, [esp+4];lpValue
	.If eax
		mov ecx, DWord ptr [eax]
		bswap ecx
		mov DWord ptr [eax], ecx 
	.endif	
	ret 4
SwapByteOrder32 endp

;#########################################################################
;# 		Input: 	Zeiger auf eine 64Bit Variable
;# 		Output:	nichts
;#########################################################################
Align sizeof(DWord)
SwapByteOrder64	proc	lpValue:ptr QWORD
	
	mov eax, [esp+4];lpValue
	.If eax
		mov ecx, DWord ptr [eax]
		mov edx, DWord ptr [eax+4]
		bswap ecx
		mov DWord ptr [eax+4], ecx 
		bswap edx
		mov DWord ptr [eax], edx 
	.endif	
	ret 4
SwapByteOrder64 endp

;#########################################################################
;# 		Input: 	Zeiger auf eine Variable mit dem Zeiger auf das SafeArray
;# 		Output:	0 oder ERROR_INVALID_VALUE
;#########################################################################
Align sizeof(DWord)
SwapByteOrderArray	proc uses ebx edi	lpValue:ptr ptr SAFEARRAY
	push ebx ;EBX muss gerettet werden
	
	mov eax, ERROR_INVALID_PARAMETER
	mov ebx, [esp+8]	; lpValue
	.If ebx && dword ptr [ebx] ; Prüfung ob wir ein Zeiger auf ein Array haben 
		mov ebx, [ebx]
		mov eax, ERROR_INVALID_PARAMETER
		; Nur Arrays mit festgelegtem Variablentyp
		.if [ebx].SAFEARRAY.fFeatures & 80h ;FADF_HAVEVARTYPE
			; Berechnung der Anzahl an belegten Bytes, jede Dimension muss multipliziert werden
			movzx eax, [ebx].SAFEARRAY.cDims
			mov edx, 1
			lea ecx, [ebx].SAFEARRAY.rgsabound
			.repeat
				imul edx, [ecx].SAFEARRAYBOUND.cElements
				add ecx, sizeof SAFEARRAYBOUND
				dec eax
			.until !eax
			; in EAX steht nun die Anzahl an Elementen
			xchg eax,edx
			push edi ;EDI muss gerettet werden
			; ermitteln und auswerten des Variablentyps
			mov ecx, [ebx-4]
			mov edi, [ebx].SAFEARRAY.pvData
			xchg ebx, eax
			.if ecx == VT_I2 ; in VB ist das ein Integer
				.repeat
					;invoke SwapByteOrder16,edi
					mov cx, word ptr [edi]
					xchg ch, cl
					mov word ptr [edi], cx
					add edi, 2
					dec ebx
				.until !ebx
				xor eax, eax
			.elseif ecx == VT_I4 || ecx == VT_R4 ; Long und Single in VB
				.repeat
					;invoke SwapByteOrder32,edi
					mov ecx, [edi]
					bswap ecx
					mov [edi], ecx
					add edi, 4
					dec ebx
				.until !ebx
				xor eax, eax
			.elseif ecx == VT_R8 || ecx == VT_CY ; Double und Currency in VB
				.repeat
					;invoke SwapByteOrder64,edi
					mov ecx, [edi]
					mov edx, [edi+4]
					bswap ecx
					mov [edi+4], ecx 
					bswap edx
					mov [edi], edx 
					add edi, 8
					dec ebx
				.until !ebx
				xor eax, eax
			.else ; leider kein unterstützter Typ
				mov eax, ERROR_INVALID_PARAMETER
			.endif
			pop edi
		.endif
	.endif	
	pop ebx
	ret 4
SwapByteOrderArray endp
OPTION EPILOGUE:EpilogueDef
OPTION PROLOGUE:PrologueDef

End DllEntry
