float Script : STANDARDSGLOBAL <
    string UIWidget = "none";
    string ScriptClass = "object";
    string ScriptOrder = "standard";
    string ScriptOutput = "color";
    string Script = "Technique=Main;";
> = 0.8;

//// UN-TWEAKABLES - AUTOMATICALLY-TRACKED TRANSFORMS ////////////////

float4x4 WvpXf : WorldViewProjection < string UIWidget="None"; >;
float4x4 WorldXf : World < string UIWidget="None"; >;
//float4x4 ViewIXf : ViewInverse < string UIWidget="None"; >;

float Timer : Time < string UIWidget = "none"; >;
float4 lightDir : Direction
<
	string Object = "DirectionalLight";
    string Space = "World";
> = {0.58f, -0.58f, 0.58f, 0.0f};
//////////////// TEXTURES ///////////////////

texture NormalTexture  <
    string ResourceName = "waves2.dds";
    string UIName =  "Normal Map";
    string ResourceType = "2D";
>;

sampler2D NormalSampler = sampler_state {
    Texture = <NormalTexture>;
    MinFilter = Linear;
    MipFilter = Linear;
    MagFilter = Linear;
    AddressU = Wrap;
    AddressV = Wrap;
};

///////// TWEAKABLE PARAMETERS //////////////////
float gFogStart = 1.0f;
float gFogRange = 200.0f;
float4 gFogColor = {0.5f, 0.5f, 0.5f, 1.0f};
bool gbRenderFog;
uniform extern float3   gEyePosW : CameraPosition;

float BumpScale <
    string UIWidget = "slider";
    float UIMin = 0.0;
    float UIMax = 5.0;
    float UIStep = 0.01;
    string UIName = "Bump Height";
> = 5.0;

float Shininess <
	string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 50.0;
    float UIStep = 0.5;
    string UIName = "Shininess";
> = 10.0;

float TexReptX <
    string UIName = "Texture Repeat X";
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 16.0;
    float UIStep = 0.1;
> = 8.0;

float TexReptY <
    string UIName = "Texture Repeat Y";
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 16.0;
    float UIStep = 0.1;
> = 4.0;

float BumpSpeedX <
    string UIName = "Bump Speed X";
    string UIWidget = "slider";
    float UIMin = -0.2;
    float UIMax = 0.2;
    float UIStep = 0.001;
> = -0.01;

float BumpSpeedY <
    string UIName = "Bump Speed Y";
    string UIWidget = "slider";
    float UIMin = -0.2;
    float UIMax = 0.2;
    float UIStep = 0.001;
> = 0.00;

float3 SpecularColor <
	string UIName = "Specular Color";
	string UIWidget = "Color";
> = {1.0f, 1.0f, 1.0f};

static float2 TextureScale = float2(TexReptX,TexReptY);
static float2 BumpSpeed = float2(BumpSpeedX,BumpSpeedY);

float3 DeepColor <
    string UIName = "Deep Water";
    string UIWidget = "Color";
> = {0.0f, 0.0f, 0.1f};

float3 ShallowColor <
    string UIName = "Shallow Water";
    string UIWidget = "Color";
> = {0.0f, 0.5f, 0.5f};

float KWater <
    string UIName = "Water Color Strength";
    string UIWidget = "slider";    
    float UIMin = 0.0;
    float UIMax = 2.0;
    float UIStep = 0.01;    
> = 1.0f;

float WaterAlpha <
    string UIName = "Water Alpha";
    string UIWidget = "slider";    
    float UIMin = 0.0;
    float UIMax = 1.0;
    float UIStep = 0.01;    
> = 0.7f;

//////////// CONNECTOR STRUCTS //////////////////

struct AppData {
    float4 Position : POSITION;   // in object space
    float2 UV : TEXCOORD0;
    //float3 Tangent  : TEXCOORD1;
    //float3 Binormal : TEXCOORD2;
    //float3 Normal   : NORMAL;
};

struct OceanVertOut {
    float4 HPosition  : POSITION;  // in clip space
    //float2 UV  : TEXCOORD0;
    float fogLerpParam : TEXCOORD0;
    float3 T2WXf1 : TEXCOORD1; // first row of the 3x3 transform from tangent to cube space
    float3 T2WXf2 : TEXCOORD2; // second row of the 3x3 transform from tangent to cube space
    float3 T2WXf3 : TEXCOORD3; // third row of the 3x3 transform from tangent to cube space
    float2 bumpUV0 : TEXCOORD4;
    float2 bumpUV1 : TEXCOORD5;
    float2 bumpUV2 : TEXCOORD6;
    float3 WorldView  : TEXCOORD7;
    
};

float cycle;

///////// SHADER FUNCTIONS ///////////////

OceanVertOut OceanVS(AppData IN)
{
    OceanVertOut OUT = (OceanVertOut)0;
    
    float4 Po = float4(IN.Position.xyz,1.0);
    // sum waves	
    Po.y = 0.0;
    float ddx = 0.0, ddy = 0.0;
    
    // compute tangent basis
    float3 B = float3(1, ddx, 0);
    float3 T = float3(0, ddy, 1);
    float3 N = float3(-ddx, 1, -ddy);
    
    float4 posH = mul(Po,WvpXf);
    OUT.HPosition = posH;
    // pass texture coordinates for fetching the normal map
    //OUT.UV = IN.UV.xy*TextureScale; 
    OUT.bumpUV0.xy = IN.UV.xy*TextureScale + cycle*BumpSpeed;
    OUT.bumpUV1.xy = IN.UV.xy*TextureScale*2.0 + cycle*BumpSpeed*4.0;
    OUT.bumpUV2.xy = IN.UV.xy*TextureScale*4.0 + cycle*BumpSpeed*8.0;

    // compute the 3x3 tranform from tangent space to object space
    float3x3 objToTangentSpace;
    // first rows are the tangent and binormal scaled by the bump scale
    objToTangentSpace[0] = BumpScale * normalize(T);
    objToTangentSpace[1] = BumpScale * normalize(B);
    objToTangentSpace[2] = normalize(N);

    OUT.T2WXf1.xyz = mul(objToTangentSpace,WorldXf[0].xyz);
    OUT.T2WXf2.xyz = mul(objToTangentSpace,WorldXf[1].xyz);
    OUT.T2WXf3.xyz = mul(objToTangentSpace,WorldXf[2].xyz);
    
    //OUT.T2WXf1 = mul(OUT.T2WXf1, 1.4f);
    //OUT.T2WXf2 = mul(OUT.T2WXf2, 1.2f); 
    //OUT.T2WXf3 = mul(OUT.T2WXf3, 1.5f);
    
    // compute the eye vector (going from shaded point to eye) in cube space
    float3 Pw = mul(Po,WorldXf).xyz;
    
    if (gbRenderFog)
    {
		float dist = distance(IN.Position, gEyePosW);
		OUT.fogLerpParam = saturate((dist - gFogStart) / gFogRange);
    }
    
    //OUT.WorldView = ViewIXf[3].xyz - Pw; // view inv. transpose contains eye position in world space in last row
    OUT.WorldView = gEyePosW - Pw;
    
    return OUT;
}


// Pixel Shaders

float4 OceanPS(OceanVertOut IN) : COLOR
{
    // sum normal maps
    float4 t0 = tex2D(NormalSampler, IN.bumpUV0)*2.0-1.0;
    float4 t1 = tex2D(NormalSampler, IN.bumpUV1)*2.0-1.0;
    float4 t2 = tex2D(NormalSampler, IN.bumpUV2)*2.0-1.0;
    float3 Nt = t0.xyz + t1.xyz + t2.xyz;
    
	float3 Vn = normalize(IN.WorldView);
	
	float3 toEye = normalize(Nt);
	float3 r = reflect(lightDir, Vn);
	float t = pow(max(dot(r, toEye), 0.0f), Shininess);
		
    float3x3 m; 
    m[0] = IN.T2WXf1;
    m[1] = IN.T2WXf2;
    m[2] = IN.T2WXf3;
    float3 Nw = mul(m,Nt);
    float3 Nn = normalize(Nw);

    float facing = 1 - max(dot(Vn, Nn), 0);

    float3 waterColor = KWater * lerp(DeepColor, ShallowColor, facing) + (t * SpecularColor);
    
    float4 f4Result;
    if (gbRenderFog)
    {
		f4Result = float4( lerp(waterColor.rgb, gFogColor, IN.fogLerpParam),WaterAlpha);
    }
    else
    {
		f4Result = float4(waterColor.rgb,WaterAlpha);
    }
    return f4Result;    
    //return float4(waterColor.rgb,WaterAlpha);
}

//////////////////// TECHNIQUE ////////////////

technique Main < string Script = "Pass=p0;"; > {
    pass p0 < string Script = "Draw=geometry;"; > {
    //cycle = abs(fmod(Timer, 200.0)-100.0f);
	VertexShader = compile vs_1_1 OceanVS();
		//ZEnable = true;
		//ZWriteEnable = true;
		//ZFunc = LessEqual;
		CullMode =ccw;
	PixelShader = compile ps_2_0 OceanPS();
    }
}

///////////////////////////////// eof ///
