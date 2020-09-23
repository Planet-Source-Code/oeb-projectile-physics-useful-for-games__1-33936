Attribute VB_Name = "modProjectiles"
Option Explicit

'This module is used for projectile calculations
'It is assumed that the target is at the same level as the launch point
'all mesurements are in feet, and all times are in seconds
'Angles are measured in radians (I have included a dunction (deg2rad) to
'convert from degrees to radians
'Use this for whatever you want, but if you find it usefull drop me a line
'Questions, comments and rants to oeb@cotse.com
'Please inform me if any of my calculations are wrong (some are bound to be)

Private Const GRAVITY As Double = 32.2
    'Gravity = 9.8 if dealing with metres
Private Const PI As Double = 3.14159265358979


Public Function Trajectory_Apex(ByVal p_Velocity As Double, _
                                ByVal p_LaunchAngle As Double) As Double
Trajectory_Apex = (p_Velocity ^ 2 * Sin(p_LaunchAngle) ^ 2) / (2 * GRAVITY)
        'The maximum height the projectile will reach
End Function
                                
Public Function GetRange(ByVal p_Velocity As Double, _
                         ByVal p_LaunchAngle As Double, _
                         ByVal p_TotalTime As Double) As Double
GetRange = (p_Velocity * p_TotalTime * Cos(p_LaunchAngle))
        'The maximum rang of the projectile
        'The best range will be got at a 45° launch angle
End Function

Public Function TotalTime(ByVal p_Velocity As Double, _
                          ByVal p_LaunchAngle As Double) As Double
TotalTime = (2 * p_Velocity * Sin(p_LaunchAngle)) / GRAVITY
    'The total time it takes the projectile to reach its target
End Function

Public Function Deg2Rad(ByVal p_Degrees As Double) As Double
Deg2Rad = p_Degrees / 180 * PI 'This function is for converting
                                'degrees to radians, all the calculations
                                'are done in radians
End Function

Public Function GetXPos(ByVal p_Velocity As Double, _
                        ByVal p_LaunchAngle As Double, _
                        ByVal p_Time As Double) As Double
GetXPos = (p_Velocity * Cos(p_LaunchAngle)) * p_Time
    'Gets the current X position of the projectile
    'p_Time is the time into the flight
End Function

Public Function GetYPos(ByVal p_Velocity As Double, _
                        ByVal p_LaunchAngle As Double, _
                        ByVal p_Time As Double) As Double
GetYPos = (p_Velocity * Sin(p_LaunchAngle) * p_Time) - (GRAVITY * (p_Time ^ 2) / 2)
    'Gets the current Y position of the projectile
    'p_Time is the time into the flight
End Function

Public Function GetCurrentVelocity(ByVal p_Velocity As Double, _
                                   ByVal p_LaunchAngle As Double, _
                                   ByVal p_Time As Double) As Double
GetCurrentVelocity = Sqr(p_Velocity ^ 2 - (2 * GRAVITY * p_Time * p_Velocity) * Sin(p_LaunchAngle) + (GRAVITY ^ 2 * p_Time ^ 2))
    'This gets the velocity of the object at any current time
End Function

