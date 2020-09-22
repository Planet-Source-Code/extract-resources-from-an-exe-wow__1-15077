VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPEInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' MSDOS File header
'
Dim FileName As String
Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

'
' IMAGE_FILE_HEADER.Characteristics
'

Const IMAGE_FILE_RELOCS_STRIPPED = &H1              ' Relocation fo stripped from file.
Const IMAGE_FILE_EXECUTABLE_IMAGE = &H2             ' File is executable  (i.e. no unresolved externel references).
Const IMAGE_FILE_LE_NUMS_STRIPPED = &H4           ' Le nunbers stripped from file.
Const IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8          ' Local symbols stripped from file.
Const IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10           ' Agressively trim workg set
Const IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20         ' App can handle >2gb addresses
Const IMAGE_FILE_BYTES_REVERSED_LO = &H80           ' Bytes of Machine word are reversed.
Const IMAGE_FILE_32BIT_Machine = &H100              ' 32 bit word Machine.
Const IMAGE_FILE_DEBUG_STRIPPED = &H200             ' Debuggg fo stripped from file  .DBG file
Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400    ' If Image is on removable media, _ copy and run from the swap file.
Const IMAGE_FILE_NET_RUN_FROM_SWAP = &H800          ' If Image is on Net, _ copy and run from the swap file.
Const IMAGE_FILE_SYSTEM = &H1000                    ' System File.
Const IMAGE_FILE_DLL = &H2000                       ' File is a DLL.
Const IMAGE_FILE_UP_SYSTEM_ONLY = &H4000            ' File should only be run on a UP Machine
Const IMAGE_FILE_BYTES_REVERSED_HI = &H8000         ' Bytes of Machine word are reversed.

'
' IMAGE_FILE_HEADER.Machine
'
Const IMAGE_FILE_MACHINE_UNKNOWN = 0
Const IMAGE_FILE_MACHINE_I386 = &H14C               ' tel 386.
Const IMAGE_FILE_MACHINE_R3000 = &H162              ' MIPS little-endian, _ =&H160 big-endian
Const IMAGE_FILE_MACHINE_R4000 = &H166              ' MIPS little-endian
Const IMAGE_FILE_MACHINE_R10000 = &H168             ' MIPS little-endian
Const IMAGE_FILE_MACHINE_WCEMIPSV2 = &H169          ' MIPS little-endian WCE v2
Const IMAGE_FILE_MACHINE_ALPHA = &H184              ' Alpha_AXP
Const IMAGE_FILE_MACHINE_POWERPC = &H1F0            ' IBM PowerPC Little-Endian
Const IMAGE_FILE_MACHINE_SH3 = &H1A2                ' SH3 little-endian
Const IMAGE_FILE_MACHINE_SH3E = &H1A4               ' SH3E little-endian
Const IMAGE_FILE_MACHINE_SH4 = &H1A6                ' SH4 little-endian
Const IMAGE_FILE_MACHINE_ARM = &H1C0                ' ARM Little-Endian

'
' Directory format.
'

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

'
' Optional header format.
'

Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16

Private Type IMAGE_OPTIONAL_HEADER
    '
    ' Standard fields.
    '

    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUnitializedData As Long
    AddressOfEntryPot As Long
    BaseOfCode As Long
    BaseOfData As Long

    '
    ' NT additional fields.
    '

    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    W32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    SubSystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Const IMAGE_SUBSYSTEM_UNKNOWN = 0               ' Unknown subsystem.
Const IMAGE_SUBSYSTEM_NATIVE = 1                ' Image doesn't require a subsystem.
Const IMAGE_SUBSYSTEM_WINDOWS_GUI = 2           ' Image runs in the Windows GUI subsystem.
Const IMAGE_SUBSYSTEM_WINDOWS_CUI = 3           ' Image runs in the Windows character subsystem.
Const IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 4        ' Image runs in the Windows CE subsystem.
Const IMAGE_SUBSYSTEM_OS2_CUI = 5               ' image runs in the OS/2 character subsystem.
Const IMAGE_SUBSYSTEM_POSIX_CUI = 7             ' image runs in the Posix character subsystem.
Const IMAGE_SUBSYSTEM_RESERVED8 = 8             ' image runs in the 8 subsystem.


'
' IMAGE_OPTIONAL_HEADER.DataDirectory indexes
'
Const IMAGE_DIRECTORY_ENTRY_EXPORT = 0          ' Export Directory
Const IMAGE_DIRECTORY_ENTRY_IMPORT = 1          ' Import Directory
Const IMAGE_DIRECTORY_ENTRY_RESOURCE = 2        ' Resource Directory
Const IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3       ' Exception Directory
Const IMAGE_DIRECTORY_ENTRY_SECURITY = 4        ' Security Directory
Const IMAGE_DIRECTORY_ENTRY_BASERELOC = 5       ' Base Relocation Table
Const IMAGE_DIRECTORY_ENTRY_DEBUG = 6           ' Debug Directory
Const IMAGE_DIRECTORY_ENTRY_COPYRIGHT = 7       ' Description String
Const IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8       ' Machine Value (MIPS GP)
Const IMAGE_DIRECTORY_ENTRY_TLS = 9             ' TLS Directory
Const IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10    ' Load Configuration Directory
Const IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11   ' Bound Import Directory in headers
Const IMAGE_DIRECTORY_ENTRY_IAT = 12            ' Import Address Table

'
' Section header
'

Const IMAGE_SIZEOF_SHORT_NAME = 8

Private Type IMAGE_SECTION_HEADER
   SecName As String * IMAGE_SIZEOF_SHORT_NAME
   PhysicalAddress As Long
   VirtualAddress  As Long
   SizeOfRawData As Long
   PointerToRawData As Long
   PointerToRelocations As Long
   PointerToLinenumbers As Long
   NumberOfRelocations As Integer
   NumberOfLinenumbers As Integer
   Characteristics  As Long
End Type

'
' IMAGE_SECTION_HEADER.Characteristics
'

Const IMAGE_SCN_TYPE_NO_PAD = &H8                       ' Reserved.
Const IMAGE_SCN_CNT_CODE = &H20                         ' Section contains code.
Const IMAGE_SCN_CNT_INITIALIZED_DATA = &H40             ' Section contains initialized data.
Const IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80           ' Section contains uninitialized data.
Const IMAGE_SCN_LNK_OTHER = &H100                       ' Reserved.
Const IMAGE_SCN_LNK_INFO = &H200                        ' Section contains comments or some other type of information.
Const IMAGE_SCN_LNK_REMOVE = &H800                      ' Section contents will not become part of image.
Const IMAGE_SCN_LNK_COMDAT = &H1000                     ' Section contents comdat.
Const IMAGE_SCN_MEM_FARDATA = &H8000
Const IMAGE_SCN_MEM_PURGEABLE = &H20000
Const IMAGE_SCN_MEM_16BIT = &H20000
Const IMAGE_SCN_MEM_LOCKED = &H40000
Const IMAGE_SCN_MEM_PRELOAD = &H80000
Const IMAGE_SCN_ALIGN_1BYTES = &H100000                 '
Const IMAGE_SCN_ALIGN_2BYTES = &H200000                 '
Const IMAGE_SCN_ALIGN_4BYTES = &H300000                 '
Const IMAGE_SCN_ALIGN_8BYTES = &H400000                 '
Const IMAGE_SCN_ALIGN_16BYTES = &H500000                ' Default alignment if no others are specified.
Const IMAGE_SCN_ALIGN_32BYTES = &H600000                '
Const IMAGE_SCN_ALIGN_64BYTES = &H700000                '
Const IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000             ' Section contains extended relocations.
Const IMAGE_SCN_MEM_DISCARDABLE = &H2000000             ' Section can be discarded.
Const IMAGE_SCN_MEM_NOT_CACHED = &H4000000              ' Section is not cachable.
Const IMAGE_SCN_MEM_NOT_PAGED = &H8000000               ' Section is not pageable.
Const IMAGE_SCN_MEM_SHARED = &H10000000                 ' Section is shareable.
Const IMAGE_SCN_MEM_EXECUTE = &H20000000                ' Section is executable.
Const IMAGE_SCN_MEM_READ = &H40000000                   ' Section is readable.
Const IMAGE_SCN_MEM_WRITE = &H80000000                  ' Section is writeable.

'
' Export directory
'

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

'
' Import modules directory entry
'

Private Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk As Long   ' RVA to original unbound IAT
    TimeDateStamp As Long        ' 0 if not bound,
                                 ' -1 if bound, and real date\time stamp
                                 '    in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                                 '    O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain As Long       ' -1 if no forwarders
    Name As Long
    FirstThunk As Long           ' RVA to IAT (if bound this IAT has actual addresses)
End Type

Private Type IMAGE_IMPORT_BY_NAME
    Hint As Integer
    FuncName(1) As Byte
End Type

Const IMAGE_ORDINAL_FLAG = &H80000000

Private Type IMAGE_BOUND_IMPORT_DESCRIPTOR
    TimeDateStamp As Long
    OffsetModuleName As Integer
    NumberOfModuleForwarderRefs As Integer
' Array of zero or more IMAGE_BOUND_FORWARDER_REF follows
End Type

Private Type IMAGE_BOUND_FORWARDER_REF
    TimeDateStamp As Long
    OffsetModuleName As Integer
    Reserved As Integer
End Type

'
' LOADED_IMAGE
'

Private Type LIST_ENTRY
   Flk As Long
   Blk As Long
End Type

Private Type LOADED_IMAGE
    ModuleName As String
    hFile As Long
    MappedAddress As Long    ' UCHAR *
    FileHeader As Long       ' IMAGE_NT_HEADERS *
    LastRvaSection As Long   ' IMAGE_SECTION_HEADER *
    NumberOfSections As Long
    Sections As Long         ' IMAGE_SECTION_HEADER *
    Characteristics As Long
    fSystemImage As Byte
    fDOSImage As Byte
    Lks As LIST_ENTRY
    SizeOfImage As Long
End Type

Private Declare Function ImageDirectoryEntryToData Lib "imagehlp" ( _
    ByVal Base As Long, _
    ByVal MappedAsImage As Byte, _
    ByVal DirectoryEntry As Integer, _
    Size As Long) As Long ' void *

Private Declare Function ImageRvaToVa Lib "imagehlp" ( _
    NtHeaders As Any, _
    ByVal Base As Long, _
    ByVal RVA As Long, _
    Optional LastRvaSection As Long) As Long ' void *

Private Declare Function MapAndLoad Lib "imagehlp" ( _
    ByVal ImageName As String, _
    ByVal DllPath As String, _
    LoadedImage As LOADED_IMAGE, _
    ByVal DotDll As Long, _
    ByVal ReadOnly As Long) As Long

Private Declare Function UnMapAndLoad Lib "imagehlp" ( _
   LoadedImage As LOADED_IMAGE) As Long

Private Declare Function CheckSumMappedFile Lib "imagehlp" ( _
   ByVal BaseAddress As Long, _
   ByVal FileLength As Long, _
   HeaderSum As Long, _
   CheckSum As Long) As Long

'
' UnDecorateSymbolName Flags
'

Const UNDNAME_COMPLETE = &H0                    ' Enable full undecoration
Const UNDNAME_NO_LEADG_UNDERSCORES = &H1      ' Remove leadg underscores from MS extended keywords
Const UNDNAME_NO_MS_KEYWORDS = &H2              ' Disable expansion of MS extended keywords
Const UNDNAME_NO_FUNCTION_RETURNS = &H4         ' Disable expansion of return type for primary declaration
Const UNDNAME_NO_ALLOCATION_MODEL = &H8         ' Disable expansion of the declaration model
Const UNDNAME_NO_ALLOCATION_LANGUAGE = &H10     ' Disable expansion of the declaration language specifier
Const UNDNAME_NO_MS_THISTYPE = &H20             ' NYI Disable expansion of MS keywords on the 'this' type for primary declaration
Const UNDNAME_NO_CV_THISTYPE = &H40             ' NYI Disable expansion of CV modifiers on the 'this' type for primary declaration
Const UNDNAME_NO_THISTYPE = &H60                ' Disable all modifiers on the 'this' type
Const UNDNAME_NO_ACCESS_SPECIFIERS = &H80       ' Disable expansion of access specifiers for members
Const UNDNAME_NO_THROW_SIGNATURES = &H100       ' Disable expansion of 'throw-signatures' for functions and poters to functions
Const UNDNAME_NO_MEMBER_TYPE = &H200            ' Disable expansion of 'static' or 'virtual'ness of members
Const UNDNAME_NO_RETURN_UDT_MODEL = &H400       ' Disable expansion of MS model for UDT returns
Const UNDNAME_32_BIT_DECODE = &H800             ' Undecorate 32-bit decorated names
Const UNDNAME_NAME_ONLY = &H1000                ' Crack only the name for primary declaration;
                                                                                                  '  return just [scope::]name.  Does expand template params
Const UNDNAME_NO_ARGUMENTS = &H2000             ' Don't undecorate arguments to function
Const UNDNAME_NO_SPECIAL_SYMS = &H4000          ' Don't undecorate special names (v-table, _ vcall, _ vector xxx, _ metatype, _ etc)

Private Declare Function UnDecorateSymbolName Lib "imagehlp" ( _
    ByVal DecoratedName As String, _
    ByVal UnDecoratedName As String, _
    ByVal UndecoratedLength As Long, _
    ByVal Flags As Long) As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal pString As Any) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal pString1 As Any, ByVal pString2 As Any) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)

Dim m_LI As LOADED_IMAGE
Dim m_IsLoaded As Boolean
Dim m_NTHdr As IMAGE_NT_HEADERS

Public Property Get CheckSum() As Long

   CheckSum = m_NTHdr.OptionalHeader.CheckSum
   
End Property

Public Function RealCheckSum() As Long
Dim HeaderSum As Long
   
   CheckSumMappedFile m_LI.MappedAddress, m_LI.SizeOfImage, HeaderSum, RealCheckSum
   
End Function



Public Property Get PreferredBase() As Long

   PreferredBase = m_NTHdr.OptionalHeader.ImageBase
   
End Property


Private Function pvGetCharacteristicsString(ByVal lChars As Long) As String

   If lChars And IMAGE_SCN_CNT_CODE Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Contains executable code/"
   If lChars And IMAGE_SCN_CNT_INITIALIZED_DATA Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Contains Initialized Data/"
   If lChars And IMAGE_SCN_CNT_UNINITIALIZED_DATA Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Contains Uninitialized Data/"
   
   If lChars And IMAGE_SCN_LNK_COMDAT Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Contains COMDAT data/"
   
   If lChars And IMAGE_SCN_ALIGN_16BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 16 bytes/"
   If lChars And IMAGE_SCN_ALIGN_1BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 1 byte/"
   If lChars And IMAGE_SCN_ALIGN_2BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 2 bytes/"
   If lChars And IMAGE_SCN_ALIGN_32BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 32 bytes/"
   If lChars And IMAGE_SCN_ALIGN_4BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 4 bytes/"
   If lChars And IMAGE_SCN_ALIGN_8BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 8 bytes/"
   If lChars And IMAGE_SCN_ALIGN_8BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 8 bytes/"
   If lChars And IMAGE_SCN_ALIGN_64BYTES Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Align 64 bytes/"
   
   If lChars And IMAGE_SCN_LNK_NRELOC_OVFL Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Extended relocations/"
   
   If lChars And IMAGE_SCN_MEM_DISCARDABLE Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Discardable/"
   If lChars And IMAGE_SCN_MEM_NOT_CACHED Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Not cached/"
   If lChars And IMAGE_SCN_MEM_NOT_PAGED Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Not paged/"
   If lChars And IMAGE_SCN_MEM_SHARED Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Shared/"
   If lChars And IMAGE_SCN_MEM_EXECUTE Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Can execute/"
   If lChars And IMAGE_SCN_MEM_READ Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Can be read/"
   If lChars And IMAGE_SCN_MEM_WRITE Then pvGetCharacteristicsString = pvGetCharacteristicsString & "Can be written/"
   
   pvGetCharacteristicsString = Left$(pvGetCharacteristicsString, Len(pvGetCharacteristicsString) - 1)
   
End Function


'
' Returns a string given its relative virtual address (RVA)
'
Private Function pvStringFromRVA(ByVal RVA As Long) As String
Dim lVA As Long

   ' Obtain the VA from RVA address
   lVA = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, RVA)

   ' Initialize return value
   pvStringFromRVA = String$(lstrlenA(lVA), 0)
   
   ' Copy the string
   lstrcpyA pvStringFromRVA, lVA
   
End Function

'
' Returns a string given its relative virtual address (RVA)
'
Private Function pvStringFromPtr(ByVal RVA As Long) As String
Dim lVA As Long

   ' Initialize return value
   pvStringFromPtr = String$(lstrlenA(RVA), 0)
   
   ' Copy the string
   lstrcpyA pvStringFromPtr, RVA
   
End Function


Public Sub SaveData(ByVal Ptr As Long, ByVal Size As Long)
Dim aData() As Byte

   ReDim aData(1 To Size)
   
   MoveMemory aData(1), ByVal Ptr, Size
   
   Open "c:\unpak\xx" For Binary As #1
      Put #1, , aData
   Close

End Sub

Public Sub EnumerateSections(aSections() As String)
Dim Idx As Long, Sect As IMAGE_SECTION_HEADER
Dim lSectPtr As Long, lLRVA As Long

   If Not m_IsLoaded Then Exit Sub
   
   Erase aSections
   ReDim aSections(0 To m_LI.NumberOfSections - 1, 2)
   
   For Idx = 0 To m_LI.NumberOfSections - 1
      
      lSectPtr = m_LI.Sections + (Idx * Len(Sect))
      
      ' Copy the section header
      MoveMemory Sect, ByVal lSectPtr, Len(Sect)
      aSections(Idx, 0) = Sect.SecName
      aSections(Idx, 1) = pvGetCharacteristicsString(Sect.Characteristics)
      aSections(Idx, 2) = Format$(Sect.SizeOfRawData, "#,##0")
      
   Next
   
End Sub


Public Sub EnumerateExportedFunctions(aFunctions() As String, lReturned As Long)
Dim ExpDir As IMAGE_EXPORT_DIRECTORY
Dim lNamesPtr As Long, lOrdinalsPtr As Long
Dim lExpPtr As Long, lSize As Long, lIdx As Long
Dim lNamePtr As Long, lNameLen As Long, lOrdinal As Integer
Dim sUndecorated As String

   lReturned = 0
   
   ' Reinitialize the array
   Erase aFunctions
    
   If Not m_IsLoaded Then Exit Sub
    
   ' Get pointer to export directory data.
   ' If the directory does not exists
   ' ImageDirectoryEntryToData returns 0.
   lExpPtr = ImageDirectoryEntryToData(m_LI.MappedAddress, 0, IMAGE_DIRECTORY_ENTRY_EXPORT, lSize)
   
   ' Check if the pointer is valid
   If lExpPtr Then
      
      ' Copy the directory from the
      ' pointer to ExpDir
      MoveMemory ExpDir, ByVal lExpPtr, Len(ExpDir)
   
      ' Return number of exported functions
      lReturned = ExpDir.NumberOfFunctions
      
      ' Redim the array
      ReDim aFunctions(0 To lReturned - 1, 0 To 3)
      
      ' Fill all items with
      ' the ordinal number
      For lIdx = 0 To lReturned - 1
         aFunctions(lIdx, 1) = Format(lIdx + ExpDir.Base, "00000") & " (&H" & Hex$(lIdx + ExpDir.Base) & ")"
      Next
      
      ' Get the virtual address to
      ' name pointers list.
      lNamesPtr = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, ExpDir.AddressOfNames)
            
      ' Get the virtual address to
      ' ordinal list.
      lOrdinalsPtr = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, ExpDir.AddressOfNameOrdinals)
                        
                           
      ' Copy all names to the array
      For lIdx = 0 To ExpDir.NumberOfNames - 1
      
         ' Get ordinal
         MoveMemory lOrdinal, ByVal lOrdinalsPtr + (lIdx * 2), Len(lOrdinal)
         
         ' Get relative pointer to name
         MoveMemory lNamePtr, ByVal lNamesPtr + (lIdx * 4), Len(lNamesPtr)
              
         ' Get the name
         aFunctions(lOrdinal, 0) = pvStringFromRVA(lNamePtr)

         ' Undecorate function name
         sUndecorated = String$(260, 0)
         If UnDecorateSymbolName(aFunctions(lOrdinal, 0), sUndecorated, 260, UNDNAME_COMPLETE) Then
            aFunctions(lOrdinal, 0) = Left$(sUndecorated, InStr(sUndecorated, vbNullChar))
         End If
         
         ' Copy ordinal to array
         ' ExpDir.Base is the base
         ' ordinal.
         aFunctions(lOrdinal, 1) = Format(lOrdinal + ExpDir.Base, "00000") & " (&H" & Hex$(lOrdinal + ExpDir.Base) & ")"
         
         ' Hint is just the index in
         ' the table
         aFunctions(lOrdinal, 2) = lIdx
         
      Next
      
   End If
   
End Sub



Public Sub EnumerateImportedModules(aModules() As String, lReturned As Long)
Dim ImpDir As IMAGE_IMPORT_DESCRIPTOR
Dim lNamePtr As Long, lIdx As Long
Dim lImpPtr As Long, lSize As Long

   lReturned = 0
   
   ' Reinitialize the array
   Erase aModules
    
   If Not m_IsLoaded Then Exit Sub
    
   ' Get pointer to import directory data.
   lImpPtr = ImageDirectoryEntryToData(m_LI.MappedAddress, 0, IMAGE_DIRECTORY_ENTRY_IMPORT, lSize)
   
   ' Check if the pointer is valid
   If lImpPtr Then
      
      ' The import modules directory
      ' is just an array that ends
      ' with an empty structure.
      ' We have to read it until we
      ' found a null module name pointer
      
      ' Copy the first entry
      ' from the pointer
      MoveMemory ImpDir, ByVal lImpPtr, Len(ImpDir)
   
      ' Read entries until Name = 0
      Do Until ImpDir.Name = 0
      
         ' Redim the array
         ReDim Preserve aModules(0 To lIdx)
         
         ' Get the module name
         aModules(lIdx) = pvStringFromRVA(ImpDir.Name)
         
         ' Increment array index
         lIdx = lIdx + 1
      
         ' Copy the next entry
         MoveMemory ImpDir, ByVal lImpPtr + (Len(ImpDir) * lIdx), Len(ImpDir)
         
      Loop
      
      lReturned = lIdx
      
   End If
   
End Sub
Public Sub EnumerateImportedFunctions(ByVal Module As String, aFunctions() As String, aOrdinals() As Integer, lReturned As Long)
Dim ImpDir As IMAGE_IMPORT_DESCRIPTOR
Dim lImpPtr As Long, lSize As Long
Dim lListPtr As Long, lIdx As Long
Dim lEntry As Long, bIsBound As Boolean

   lReturned = 0
   
   ' Reinitialize the array
   Erase aFunctions
    
   If Not m_IsLoaded Then Exit Sub
    
   ' Get pointer to import directory data.
   lImpPtr = ImageDirectoryEntryToData(m_LI.MappedAddress, 0, IMAGE_DIRECTORY_ENTRY_IMPORT, lSize)
   
   ' Check if the pointer is valid
   If lImpPtr Then
            
      ' Copy the first entry
      ' from the pointer
      MoveMemory ImpDir, ByVal lImpPtr, Len(ImpDir)
   
      ' Read all entries until we
      ' found the module name
      Do Until ImpDir.Name = 0 Or UCase$(pvStringFromRVA(ImpDir.Name)) = UCase$(Module)
      
         ' Move the pointer to next entry
         lImpPtr = lImpPtr + Len(ImpDir)
         
         ' Get the next entry
         MoveMemory ImpDir, ByVal lImpPtr, Len(ImpDir)
      
      Loop
           
      ' If ImpDir.ModuleName = 0 the
      ' module is not imported
            
      If ImpDir.Name <> 0 Then
   
         If ImpDir.TimeDateStamp = -1 Then
                     
            lListPtr = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, ImpDir.OriginalFirstThunk)
               
         Else
         
            ' Get pointer to the list
            lListPtr = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, ImpDir.FirstThunk)
            
         End If
         
         ' Get first entry
         MoveMemory lEntry, ByVal lListPtr, 4
         
         Do Until lEntry = 0
         
            ReDim Preserve aFunctions(0 To lIdx)
            ReDim Preserve aOrdinals(0 To lIdx)
                        
            ' Check if the function is
            ' imported by ordinal
            If (lEntry And IMAGE_ORDINAL_FLAG) = IMAGE_ORDINAL_FLAG Then
            
               ' Add the ordinal
               aOrdinals(lIdx) = lEntry And &H7FFF&
               
               If (lEntry And &H8000&) Then aOrdinals(lIdx) = aOrdinals(lIdx) Or &H8000
         
            Else
               
               ' The entry points to a
               ' IMAGE_IMPORT_BY_NAME struct
               
               ' Get the function name.
               aFunctions(lIdx) = pvStringFromRVA(lEntry + 2)
               
               ' Convert RVA to VA
               lEntry = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, lEntry)
               
               ' Copy the Hint from pointer
               MoveMemory aOrdinals(lIdx), ByVal lEntry, 2
               
            End If
            
            ' Move pointer to next item
            lListPtr = lListPtr + 4
            
            ' Get next entry
            MoveMemory lEntry, ByVal lListPtr, 4
            
            ' Increment array index
            lIdx = lIdx + 1
            
         Loop
      
      End If
      
      lReturned = lIdx
      
   End If
   
End Sub

'
' Return the size of the image in memory
'
Public Property Get ImageSize() As Long

   ImageSize = m_LI.SizeOfImage
   
End Property


Public Property Get ImageVer() As String

   ImageVer = m_NTHdr.OptionalHeader.MajorImageVersion & "." & m_NTHdr.OptionalHeader.MinorImageVersion
   
End Property

Public Function Machine() As Long

   Machine = m_NTHdr.FileHeader.Machine
   
End Function

Public Function MachineName() As String
 
   Select Case m_NTHdr.FileHeader.Machine
   
      Case IMAGE_FILE_MACHINE_I386
         MachineName = "x86"

      Case IMAGE_FILE_MACHINE_R3000, IMAGE_FILE_MACHINE_R4000, IMAGE_FILE_MACHINE_R10000, IMAGE_FILE_MACHINE_WCEMIPSV2
         MachineName = "MIPS"

      Case IMAGE_FILE_MACHINE_ALPHA
         MachineName = "Alpha"
      
      Case IMAGE_FILE_MACHINE_POWERPC
         MachineName = "PowerPC"
         
      Case IMAGE_FILE_MACHINE_SH3
         MachineName = "SH3"
      
      Case IMAGE_FILE_MACHINE_SH3E
         MachineName = "SH3E"
      
      Case IMAGE_FILE_MACHINE_SH4
         MachineName = "SH4"
         
      Case IMAGE_FILE_MACHINE_ARM
         MachineName = "ARM"
         
      Case Else
         MachineName = "Unknown"
         
   End Select
   
End Function


Public Property Get SubSystem() As String

   Select Case m_NTHdr.OptionalHeader.SubSystem
      Case IMAGE_SUBSYSTEM_NATIVE
         SubSystem = "Native"
      Case IMAGE_SUBSYSTEM_WINDOWS_GUI
         SubSystem = "Windows GUI"
      Case IMAGE_SUBSYSTEM_WINDOWS_CUI
         SubSystem = "Windowns Console"
      Case IMAGE_SUBSYSTEM_WINDOWS_CE_GUI
         SubSystem = "Windown CE"
      Case IMAGE_SUBSYSTEM_OS2_CUI
         SubSystem = "OS/2 Console"
      Case IMAGE_SUBSYSTEM_POSIX_CUI
         SubSystem = "Posix Console"
      Case Else
         SubSystem = "Unknown"
   End Select
   
End Property

Public Property Get SubSystemVer() As String

   SubSystemVer = m_NTHdr.OptionalHeader.MajorSubsystemVersion & "." & m_NTHdr.OptionalHeader.MinorSubsystemVersion
   
End Property

Public Property Get OSVer() As String

   OSVer = m_NTHdr.OptionalHeader.MajorOperatingSystemVersion & "." & m_NTHdr.OptionalHeader.MinorOperatingSystemVersion
   
End Property


Public Property Get LinkerVer() As String

   LinkerVer = m_NTHdr.OptionalHeader.MajorLinkerVersion & "." & m_NTHdr.OptionalHeader.MinorLinkerVersion
   
End Property


'
' Return the number of sections
'
Public Property Get Sections() As Long

   Sections = m_LI.NumberOfSections
   
End Property



Public Sub Load(ByVal File As String)
Dim lRet As Long

   ' Unload previous mapped file
   If m_IsLoaded Then Unload
   
   ' Load and map the file
   ' into the current address space
   lRet = MapAndLoad(File, vbNullString, m_LI, True, True)
   
   If lRet Then
      
      ' Set the loaded flag
      m_IsLoaded = True
      
      ' Copy the NT header from pointer
      MoveMemory m_NTHdr, ByVal m_LI.FileHeader, Len(m_NTHdr)
      
   Else
   
      Err.Raise Err.LastDllError, , "Cannot load and map the file. GetLastError = " & Err.LastDllError
      
   End If
   
End Sub


Public Sub Unload()

   ' Unload any loaded file
   If m_IsLoaded Then
      
      UnMapAndLoad m_LI
      
      m_IsLoaded = False
      
   End If
   
End Sub


Private Sub Class_Terminate()

   Unload
   
End Sub




