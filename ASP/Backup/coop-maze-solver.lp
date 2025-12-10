% Name: Co-op Maze Running Solver

% KRR Methodology: Typical planning problem.
%                  Goes through N number of steps and tries every allowed action

% Output: occurs(Action, TimeStep).

% Input:
% cell(X, Y). % (X,Y) is a traversable cell
% player(X). % (X) is a player
% wall(X, Y). % (X,Y) is a wall, cannot be entered
% door(X, Y). % (X,Y) is a door, blocks movement when closed
% trap(X, Y). % (X,Y) is a freeze trap
% lever(X, Y, Z, W). % lever at (X,Y) controls door/trap at (Z,W)
% adjacent(X, Y, Z, W). % (X,Y) is adjacent to (Z,W)
% exit(X, Y). % exit cell; goal if any runner reaches it

% --------------------------
% Objects and their types
% --------------------------
% 2 Players:
player(1).
player(2).
% 2 Player Starts:
holds(at(1, 2, 2), 0).
holds(at(2, 13, 2), 0).
% 101 Cells:
cell(2,2).
cell(5,2).
cell(6,2).
cell(7,2).
cell(8,2).
cell(9,2).
cell(10,2).
cell(11,2).
cell(12,2).
cell(13,2).
cell(14,2).
cell(2,3).
cell(5,3).
cell(6,3).
cell(7,3).
cell(8,3).
cell(9,3).
cell(10,3).
cell(11,3).
cell(12,3).
cell(13,3).
cell(14,3).
cell(2,4).
cell(14,4).
cell(2,5).
cell(10,5).
cell(11,5).
cell(12,5).
cell(14,5).
cell(2,6).
cell(3,6).
cell(4,6).
cell(5,6).
cell(6,6).
cell(7,6).
cell(10,6).
cell(11,6).
cell(12,6).
cell(14,6).
cell(2,7).
cell(4,7).
cell(5,7).
cell(6,7).
cell(7,7).
cell(14,7).
cell(2,8).
cell(4,8).
cell(7,8).
cell(9,8).
cell(11,8).
cell(12,8).
cell(13,8).
cell(14,8).
cell(2,9).
cell(4,9).
cell(5,9).
cell(7,9).
cell(9,9).
cell(14,9).
cell(2,10).
cell(4,10).
cell(5,10).
cell(7,10).
cell(9,10).
cell(10,10).
cell(11,10).
cell(12,10).
cell(14,10).
cell(2,11).
cell(4,11).
cell(5,11).
cell(7,11).
cell(9,11).
cell(10,11).
cell(11,11).
cell(12,11).
cell(14,11).
cell(2,12).
cell(4,12).
cell(5,12).
cell(7,12).
cell(9,12).
cell(10,12).
cell(11,12).
cell(12,12).
cell(14,12).
cell(2,13).
cell(7,13).
cell(14,13).
cell(2,14).
cell(3,14).
cell(4,14).
cell(5,14).
cell(7,14).
cell(8,14).
cell(9,14).
cell(10,14).
cell(11,14).
cell(12,14).
cell(13,14).
cell(14,14).

% 2 Freeze Traps:
trap(6,2).
trap(11,14).

% 1 Doors:
door(2,5).

% 124 Walls:
wall(1,1).
wall(2,1).
wall(3,1).
wall(4,1).
wall(5,1).
wall(6,1).
wall(7,1).
wall(8,1).
wall(9,1).
wall(10,1).
wall(11,1).
wall(12,1).
wall(13,1).
wall(14,1).
wall(15,1).
wall(1,2).
wall(3,2).
wall(4,2).
wall(15,2).
wall(1,3).
wall(3,3).
wall(4,3).
wall(15,3).
wall(1,4).
wall(3,4).
wall(4,4).
wall(5,4).
wall(6,4).
wall(7,4).
wall(8,4).
wall(9,4).
wall(10,4).
wall(11,4).
wall(12,4).
wall(13,4).
wall(15,4).
wall(1,5).
wall(3,5).
wall(4,5).
wall(5,5).
wall(6,5).
wall(7,5).
wall(8,5).
wall(9,5).
wall(13,5).
wall(15,5).
wall(1,6).
wall(8,6).
wall(9,6).
wall(13,6).
wall(15,6).
wall(1,7).
wall(3,7).
wall(8,7).
wall(9,7).
wall(10,7).
wall(11,7).
wall(12,7).
wall(13,7).
wall(15,7).
wall(1,8).
wall(3,8).
wall(5,8).
wall(6,8).
wall(8,8).
wall(10,8).
wall(15,8).
wall(1,9).
wall(3,9).
wall(6,9).
wall(8,9).
wall(10,9).
wall(11,9).
wall(12,9).
wall(13,9).
wall(15,9).
wall(1,10).
wall(3,10).
wall(6,10).
wall(8,10).
wall(13,10).
wall(15,10).
wall(1,11).
wall(3,11).
wall(6,11).
wall(8,11).
wall(13,11).
wall(15,11).
wall(1,12).
wall(3,12).
wall(6,12).
wall(8,12).
wall(13,12).
wall(15,12).
wall(1,13).
wall(3,13).
wall(4,13).
wall(5,13).
wall(6,13).
wall(8,13).
wall(9,13).
wall(10,13).
wall(11,13).
wall(12,13).
wall(13,13).
wall(15,13).
wall(1,14).
wall(6,14).
wall(15,14).
wall(1,15).
wall(2,15).
wall(3,15).
wall(4,15).
wall(5,15).
wall(6,15).
wall(7,15).
wall(8,15).
wall(9,15).
wall(10,15).
wall(11,15).
wall(12,15).
wall(13,15).
wall(14,15).
wall(15,15).

% 1 Exits:
exit(7,8).

% 1 Levers:
lever(14,6, 6, 2).


% --------------------------
% Fluents
% --------------------------

% player P is at (X, Y)
fluent(inertial, at(P, X, Y)) :- tile(X, Y), player(P).
% player P is frozen this turn
fluent(inertial, frozen(P)) :- player(P).
% the door at (X, Y) is open
fluent(inertial, open(X, Y)) :- door(X, Y).
% the trap at (X, Y) is deactivated
fluent(inertial, deactivated(X, Y)) :- trap(X, Y).

% Flip-flopper for moving at half speed
% freeze_turn is only true every other step
freeze_turn(I) :- not freeze_turn_prev(I), step(I).
freeze_turn_prev(I) :- freeze_turn(I-1), step(I).


% --------------------------
% Static
% --------------------------

% A tile is ANY type of coordinate object
tile(X,Y) :- cell(X,Y).
tile(X,Y) :- wall(X,Y).
tile(X,Y) :- door(X,Y).
tile(X,Y) :- trap(X,Y).
tile(X,Y) :- exit(X,Y).
tile(X,Y) :- lever(X,Y,_,_).

% Adjacency in 4 directions, but only if both coordinates exist as tiles
adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X + 1, Y1 = Y.
adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X - 1, Y1 = Y.
adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X,     Y1 = Y + 1.
adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X,     Y1 = Y - 1.

% Adjacency in the diagonals
% Disabled
% adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X + 1, Y1 = Y + 1.
% adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X + 1, Y1 = Y - 1.
% adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X - 1, Y1 = Y + 1.
% adjacent(X,Y,X1,Y1) :- tile(X,Y), tile(X1,Y1), X1 = X - 1, Y1 = Y - 1.

% --------------------------
% Actions
% --------------------------

% player P will move to (X, Y)
action(move(P, X, Y)) :- player(P), tile(X, Y).
% player P pulls the lever at (X, Y)
action(pull_lever(P, X, Y)) :- player(P), lever(X, Y, Z, W).
% player P opens the door at (X, Y)
action(open_door(P, X, Y)) :- player(P), door(X, Y).
% player P does nothing
action(wait(P)) :- player(P).

% --------------------------
% Axioms
% --------------------------

% Dynamic causal laws: direct effects of actions

% Player P will be at X, Y if they move to X, Y in the previous step
holds(at(P, X, Y), I1) :- occurs(move(P, X, Y), I),
                      I < n,
                      next(I, I1).

% Player P will no longer be at Z, W if they move to X, Y in the previous step
-holds(at(P, Z, W), I1) :- occurs(move(P, X, Y), I),
                      holds(at(P, Z, W), I),
                      I < n,
                      next(I, I1).

% A door is open if open_door occurs at it's location
holds(open(X, Y), I1) :- occurs(open_door(P, X, Y), I),
                      I < n,
                      next(I, I1).

% A door is open if pull_lever occurs on a lever that is attached to the door
holds(open(X, Y), I1) :- occurs(pull_lever(P, W, Z), I),
                      lever(W, Z, X, Y),
                      I < n,
                      next(I, I1).

% A trap is deactivated if pull_lever occurs on a lever that is attached to the trap
holds(deactivated(X, Y), I1) :- occurs(pull_lever(P, W, Z), I),
                      lever(W, Z, X, Y),
                      I < n,
                      next(I, I1).

% A player is frozen every other turn if they have entered a freeze trap that was not deactivated
holds(frozen(P), I1) :- holds(at(P, X, Y), I),
                      -holds(deactivated(X, Y), I),
                      trap(X, Y),
                      I < n,
                      next(I, I1).

% State constraints

% A player can only be in one place at a time
-holds(at(P, X, Y), I) :- holds(at(P, Z, W), I), Z != X, Y != W, tile(X, Y).
% Rule disabled because I want to allow players to stack
% -holds(at(P, X, Y), I) :- holds(at(P1, X, Y), I), P != P1, player(P).

% Executability conditions
% These were changed to constraints to improve performance

% Must be on the lever to pull it
:- occurs(pull_lever(P, X, Y), I),
  not holds(at(P, X, Y), I).

% Must be adjacent to a door to open it
:- occurs(open_door(P, X, Y), I),
  holds(at(P, Z, W), I),
  not adjacent(X, Y, Z, W).

% Cannot open an already open door
:- occurs(open_door(P, X, Y), I),
  holds(open(X, Y), I).


% Cannot move to non-adjacent tile
:- occurs(move(P,X,Y), I),
  holds(at(P,Z,W), I),
  not adjacent(Z,W,X,Y).

% Cannot move into a closed door
:- occurs(move(P,X,Y), I),
  door(X,Y),
  not holds(open(X,Y), I).

% Cannot move into a wall (just in case)
:- occurs(move(P,X,Y), I),
  wall(X,Y).

% Turning each action into an agent move to group them
agent(move(P, X, Y), P) :- player(P), tile(X, Y).
agent(pull_lever(P, X, Y), P) :- player(P), lever(X, Y, Z, W).
agent(open_door(P, X, Y), P) :- player(P), door(X, Y).
agent(wait(P), P) :- player(P).

% Agent moves without "wait" included so that the player can
% wait when they are frozen
agent_no_wait(move(P, X, Y), P) :- player(P), tile(X, Y).
agent_no_wait(pull_lever(P, X, Y), P) :- player(P), lever(X, Y, Z, W).
agent_no_wait(open_door(P, X, Y), P) :- player(P), door(X, Y).

% Two actions cannot occur by the same player in the same step
:- occurs(A1, I),
  occurs(A2, I),
  agent(A1, P),
  agent(A2, P),
  A1 != A2,
  step(I).

% The player cannot execute most actions when frozen for that turn
:- occurs(A, I),
   agent_no_wait(A, P),
   holds(frozen(P), I),
   freeze_turn(I),
   step(I).

%% GENERAL INERTIA AXIOMS FOR INERTIAL FLUENTS:

holds(F,I1) :- fluent(inertial,F),
               holds(F,I),
               not -holds(F,I1),
               I < n,
               next(I, I1).

-holds(F,I1) :- fluent(inertial,F),
		        -holds(F,I),
                not holds(F,I1),
                I < n,
                next(I, I1).

% GENERAL CWA FOR DEFINED FLUENTS

-holds(F, I) :- fluent(defined, F),
		        not holds(F, I),
		        step(I).

% --------------------------
% A particular story:
%      Initially, ...
% --------------------------

#const n=40. % change the number
step(0..n).

next(I, I1) :- step(I), step(I1), I1 = I+1.

% Initial situation: what holds at time step 0

% CWA for time step 0
-holds(F, 0) :- fluent(inertial, F), not holds(F, 0).

% --------------------------
% Planning Module
% --------------------------

% This rule changes depending on the planning problem

% Player is at the exit
player_exit(P, I) :-
    holds(at(P, X, Y), I),
    exit(X, Y).

% Player is not at the exit
player_not_exit(I) :-
    player(P),
    step(I),
    not player_exit(P, I).

% Goal is for every player to be at the exit
goal(I) :-
    step(I),
    not player_not_exit(I).

success :- goal(I), I <= n.

:- not success.

occurs(A, I) | -occurs(A, I) :- action(A),
                                step(I),
                                not goal(I),
                                I < n.

something_happened(I) :- occurs(A, I).

% Sorted based on occurs
:~ occurs(A,I). [1,I,A]

:- step(J),
  goal(I),
  not goal(I0),
  I0 = I-1,
  J < I,
  not something_happened(J).

#show occurs/2.
