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
